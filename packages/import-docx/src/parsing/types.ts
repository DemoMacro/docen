export interface ListInfo {
  type: "bullet" | "ordered";
  start?: number;
}

export type ListTypeMap = Map<string, ListInfo>;
