// 3D 选件视图：GLB 节点名 = 零件名，点击即打标。只负责「看与选」，不产出任何几何。
import * as THREE from 'three';
import { GLTFLoader } from 'three/addons/loaders/GLTFLoader.js';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { api } from './api.js';
import {
  bus, state, figureColor, figuresContaining, resolveMembers, toggleSelect, partNames,
  setSelection,
} from './state.js';

const BASE = new THREE.Color('#b9b4a9');       // 未入图：暖灰（纸上铅稿的灰）
const EDGE = new THREE.Color('#3a372f');
const SELECT_EMISSIVE = new THREE.Color('#1c4d8f');
const HOVER_EMISSIVE = new THREE.Color('#6a4f00');

let renderer; let scene; let camera; let controls; let raycaster;
const meshesByName = new Map();   // 零件名 -> Mesh[]
let hovered = null;
let isolate = false;
let flashUntil = 0;
const flashSet = new Set();
let boxMode = false;
const samples = new Map();        // 零件名 -> 世界坐标采样点 Float32Array（框选用，载入后算一次）
const SAMPLE_PER_MESH = 360;

function collectMeshes(rootObj) {
  // GLTFLoader 会给重名对象追加 _1/_2 后缀（同一零件的节点与 mesh 同名也算重名），
  // 所以 mesh 名不能直接当零件名用。以 assembly.json 的零件表为权威做归一：
  // 沿自身与祖先收集候选名，剥掉 _N 后缀，取第一个存在于零件表里的。
  const known = new Set(partNames());
  const canonical = (raw) => {
    const t = (raw || '').trim();
    if (known.has(t)) return t;
    const stripped = t.replace(/_\d+$/, '');
    return known.has(stripped) ? stripped : null;
  };
  rootObj.traverse((obj) => {
    if (!obj.isMesh) return;
    let name = null;
    for (let node = obj; node && !name; node = node.parent) name = canonical(node.name);
    if (!name) return; // 装配根或无法归一的对象不参与选件
    obj.userData.part = name;
    obj.material = new THREE.MeshStandardMaterial({
      color: BASE.clone(), metalness: 0.05, roughness: 0.82,
      emissive: 0x000000, transparent: true, opacity: 1.0,
    });
    obj.castShadow = false;
    if (!meshesByName.has(name)) meshesByName.set(name, []);
    meshesByName.get(name).push(obj);
  });
}

export function repaint() {
  const active = state.activeFigure;
  const activeMembers = active
    ? resolveMembers((state.plan.figures || []).find((f) => f.id === active) || { members: [] })
    : new Set();
  const flashing = performance.now() < flashUntil;
  meshesByName.forEach((meshes, name) => {
    const figs = figuresContaining(name);
    const inActive = activeMembers.has(name);
    const color = figs.length
      ? new THREE.Color(figureColor(inActive ? active : figs[0].id))
      : BASE.clone();
    if (figs.length && !inActive) color.lerp(BASE, 0.55); // 非当前图：褪成底色调
    meshes.forEach((m) => {
      m.visible = !isolate || inActive || state.selection.has(name);
      m.material.color.copy(color);
      m.material.opacity = figs.length || !active ? 1.0 : 0.92;
      m.material.emissive.set(0x000000);
      m.material.emissiveIntensity = 0.0;
      if (state.selection.has(name)) {
        m.material.emissive.copy(SELECT_EMISSIVE);
        m.material.emissiveIntensity = 0.55;
      }
      if (flashing && flashSet.has(name)) {
        m.material.emissive.copy(SELECT_EMISSIVE);
        m.material.emissiveIntensity = 0.9;
      }
      if (hovered === name) {
        m.material.emissive.copy(state.selection.has(name) ? SELECT_EMISSIVE : HOVER_EMISSIVE);
        m.material.emissiveIntensity = 0.75;
      }
    });
  });
}

export function flashParts(names, ms = 1600) {
  flashSet.clear();
  names.forEach((n) => flashSet.add(n));
  flashUntil = performance.now() + ms;
  repaint();
  setTimeout(repaint, ms + 30);
}

export function frameAll(padding = 1.35) {
  // 首次调用发生在第一帧渲染之前，此时 glTF 各节点的世界矩阵尚未由渲染循环
  // 更新，expandByObject 会量出一个错误的大包围盒（实测把相机推远了 9 倍）。
  scene.updateMatrixWorld(true);
  const box = new THREE.Box3();
  meshesByName.forEach((meshes) => meshes.forEach((m) => {
    if (m.visible) box.expandByObject(m);
  }));
  if (box.isEmpty()) return;
  const size = box.getSize(new THREE.Vector3()).length();
  const center = box.getCenter(new THREE.Vector3());
  controls.target.copy(center);
  const dir = new THREE.Vector3(1, -0.7, 0.8).normalize();
  camera.position.copy(center).addScaledVector(dir, size * padding);
  camera.near = size / 200; camera.far = size * 20;
  camera.updateProjectionMatrix();
  controls.update();
}

function pick(event) {
  const rect = renderer.domElement.getBoundingClientRect();
  const ndc = new THREE.Vector2(
    ((event.clientX - rect.left) / rect.width) * 2 - 1,
    -((event.clientY - rect.top) / rect.height) * 2 + 1,
  );
  raycaster.setFromCamera(ndc, camera);
  const all = [];
  meshesByName.forEach((meshes) => meshes.forEach((m) => { if (m.visible) all.push(m); }));
  const hits = raycaster.intersectObjects(all, false);
  return hits.length ? hits[0].object.userData.part : null;
}

// ---- 框选（CAD 惯例：左→右窗选，全在框内才选；右→左交叉选，碰到即选）----
function buildSamples() {
  samples.clear();
  scene.updateMatrixWorld(true);
  const v = new THREE.Vector3();
  meshesByName.forEach((meshes, name) => {
    const pts = [];
    meshes.forEach((m) => {
      const pos = m.geometry?.attributes?.position;
      if (!pos) return;
      const step = Math.max(1, Math.floor(pos.count / SAMPLE_PER_MESH));
      for (let i = 0; i < pos.count; i += step) {
        v.fromBufferAttribute(pos, i).applyMatrix4(m.matrixWorld);
        pts.push(v.x, v.y, v.z);
      }
    });
    samples.set(name, new Float32Array(pts));
  });
}

function partsInRect(r, crossing) {
  const rect = renderer.domElement.getBoundingClientRect();
  const v = new THREE.Vector3();
  const hit = [];
  samples.forEach((pts, name) => {
    const meshes = meshesByName.get(name) || [];
    if (!meshes.some((m) => m.visible) || !pts.length) return;
    let any = false; let all = true;
    for (let i = 0; i < pts.length; i += 3) {
      v.set(pts[i], pts[i + 1], pts[i + 2]).project(camera);
      if (v.z > 1) { all = false; continue; }          // 相机背后
      const x = (v.x + 1) / 2 * rect.width;
      const y = (1 - v.y) / 2 * rect.height;
      const inside = x >= r.x0 && x <= r.x1 && y >= r.y0 && y <= r.y1;
      any = any || inside;
      all = all && inside;
      if (crossing && any) break;
      if (!crossing && !all) break;
    }
    if (crossing ? any : all) hit.push(name);
  });
  return hit;
}

export function setBoxMode(on) {
  boxMode = !!on;
  if (controls) controls.enabled = !boxMode;
  const host = renderer?.domElement?.parentElement;
  if (host) host.classList.toggle('box-mode', boxMode);
  bus.emit('box-mode', boxMode);
}
export const isBoxMode = () => boxMode;

function attachMarquee(container) {
  const band = document.createElement('div');
  band.className = 'marquee';
  band.hidden = true;
  container.appendChild(band);
  let start = null;
  const local = (e) => {
    const r = renderer.domElement.getBoundingClientRect();
    return [e.clientX - r.left, e.clientY - r.top];
  };
  renderer.domElement.addEventListener('pointerdown', (e) => {
    if (!boxMode || e.button !== 0) return;
    start = local(e);
    renderer.domElement.setPointerCapture(e.pointerId);
    band.hidden = false;
    Object.assign(band.style, { left: `${start[0]}px`, top: `${start[1]}px`, width: '0px', height: '0px' });
  });
  renderer.domElement.addEventListener('pointermove', (e) => {
    if (!start) return;
    const [x, y] = local(e);
    const crossing = x < start[0];
    band.classList.toggle('crossing', crossing);
    Object.assign(band.style, {
      left: `${Math.min(x, start[0])}px`, top: `${Math.min(y, start[1])}px`,
      width: `${Math.abs(x - start[0])}px`, height: `${Math.abs(y - start[1])}px`,
    });
  });
  renderer.domElement.addEventListener('pointerup', (e) => {
    if (!start) return;
    const [x, y] = local(e);
    const r = { x0: Math.min(x, start[0]), x1: Math.max(x, start[0]),
      y0: Math.min(y, start[1]), y1: Math.max(y, start[1]) };
    const crossing = x < start[0];
    start = null;
    band.hidden = true;
    if (r.x1 - r.x0 < 4 || r.y1 - r.y0 < 4) return;   // 太小当作误触
    const hit = partsInRect(r, crossing);
    let next;
    if (e.altKey) next = [...state.selection].filter((n) => !hit.includes(n));
    else if (e.shiftKey) next = [...new Set([...state.selection, ...hit])];
    else next = hit;
    setSelection(next);
    bus.emit('box-selected', { count: hit.length, crossing });
  });
}

export async function initViewer(container) {
  renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  container.appendChild(renderer.domElement);
  scene = new THREE.Scene();
  camera = new THREE.PerspectiveCamera(40, 1, 0.1, 5000);
  controls = new OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true;
  raycaster = new THREE.Raycaster();

  scene.add(new THREE.HemisphereLight(0xfffdf5, 0x8a8577, 1.15));
  const key = new THREE.DirectionalLight(0xffffff, 1.5);
  key.position.set(1.5, -2, 2.5);
  scene.add(key);
  const rim = new THREE.DirectionalLight(0xdfe8ff, 0.5);
  rim.position.set(-2, 1.5, -1);
  scene.add(rim);

  const gltf = await new GLTFLoader().loadAsync(api.modelUrl());
  scene.add(gltf.scene);
  collectMeshes(gltf.scene);
  frameAll();
  buildSamples();
  attachMarquee(container);

  const resize = () => {
    const { clientWidth: w, clientHeight: h } = container;
    renderer.setSize(w, h, false);
    camera.aspect = w / h;
    camera.updateProjectionMatrix();
  };
  new ResizeObserver(resize).observe(container);
  resize();

  let downAt = null;
  renderer.domElement.addEventListener('pointerdown', (e) => { downAt = [e.clientX, e.clientY]; });
  renderer.domElement.addEventListener('pointerup', (e) => {
    if (!downAt || boxMode) { downAt = null; return; }
    const moved = Math.hypot(e.clientX - downAt[0], e.clientY - downAt[1]);
    downAt = null;
    if (moved > 5) return;                    // 拖转视角不算点击
    const name = pick(e);
    if (name) toggleSelect(name);
  });
  renderer.domElement.addEventListener('pointermove', (e) => {
    if (boxMode && e.buttons) return;
    const name = pick(e);
    if (name !== hovered) {
      hovered = name;
      container.style.cursor = name ? 'pointer' : 'grab';
      container.dispatchEvent(new CustomEvent('part-hover', { detail: name }));
      repaint();
    }
  });

  const animate = () => {
    requestAnimationFrame(animate);
    controls.update();
    renderer.render(scene, camera);
  };
  animate();

  bus.on('plan', repaint);
  bus.on('selection', repaint);
  bus.on('active-figure', repaint);
  repaint();   // 初次上色：viewer 在 booted 事件之后才建好，错过了那班车
  const handle = {
    setIsolate(on) { isolate = on; repaint(); frameAll(); },
    setBoxMode,
    frameAll,
    partCount: meshesByName.size,
    meshesByName, scene, camera,   // 本地调试句柄（单机工具，无隐私面）
  };
  window.planStudioViewer = handle;
  return handle;
}
