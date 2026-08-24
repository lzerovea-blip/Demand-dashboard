export type ImageViewerDirection = -1 | 1;

export function moveImageViewerIndex(currentIndex: number, imageCount: number, direction: ImageViewerDirection): number {
  if (imageCount <= 1) return 0;
  const normalizedIndex = ((currentIndex % imageCount) + imageCount) % imageCount;
  return (normalizedIndex + direction + imageCount) % imageCount;
}
