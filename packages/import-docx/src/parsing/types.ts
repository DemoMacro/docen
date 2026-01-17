export interface ListInfo {
  type: "bullet" | "ordered";
  start?: number;
}

export type ListTypeMap = Map<string, ListInfo>;

export interface ImageInfo {
  src: string; // data URL (e.g., "data:image/png;base64,...")
  width?: number;
  height?: number;
}
