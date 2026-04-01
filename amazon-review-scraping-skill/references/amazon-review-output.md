# Amazon Review Output

## JSON fields

- `asin`: product ASIN
- `host`: Amazon marketplace host, for example `www.amazon.sg`
- `productUrl`: original product page
- `reviewsUrl`: review landing page
- `fetchedAt`: ISO timestamp
- `sessionDir`: persistent Playwright session directory
- `viewResults`: counts by review view
- `mediaDir`: downloaded review-image directory
- `reviewCount`: deduplicated written review count
- `imageReviewCount`: review rows containing images
- `downloadedImageCount`: local image files saved
- `reviews[]`: per-review payload

## `reviews[]` fields

- `reviewId`
- `author`
- `title`
- `body`
- `ratingText`
- `rating`
- `dateText`
- `verifiedPurchase`
- `sourceViews`
- `images[]`

## `images[]` fields

- `imageIndex`
- `originalUrl`
- `thumbnailUrl`
- `localPath`
