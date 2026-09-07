// 3D 选件视图：GLB 节点名 = 零件名，点击即打标。只负责「看与选」，不产出任何几何。
import * as THREE from 'three';
import { GLTFLoader } from 'three/addons/loaders/GLTFLoader.js';
import { OrbitControls } from 'three/addons/controls/OrbitControls.js';
import { api } from './api.js';
import {
  bus, state, figureColor, figuresContaining, resolveMembers, toggleSelect, partNames,
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
    if (!downAt) return;
    const moved = Math.hypot(e.clientX - downAt[0], e.clientY - downAt[1]);
    downAt = null;
    if (moved > 5) return;                    // 拖转视角不算点击
    const name = pick(e);
    if (name) toggleSelect(name);
  });
  renderer.domElement.addEventListener('pointermove', (e) => {
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
  return {
    setIsolate(on) { isolate = on; repaint(); frameAll(); },
    frameAll,
    partCount: meshesByName.size,
  };
}
