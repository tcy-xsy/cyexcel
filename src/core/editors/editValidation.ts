import type { CellRect } from "../layout";
import type { DataValidationResult } from "../sheet/validations";

/**
 * 编辑器提交前的校验结论（数据验证设计稿 §7.1）。`EditorManager` 只认这个形状，
 * 不认识规则本身：规则怎么查、值怎么判是集成层（`CyExcelCore`）的事。
 *
 * - `accept`：照常提交。
 * - `reject`（停止）：不提交，编辑器、输入值和焦点都保留，弹出错误提示。
 * - `confirm`（警告）：弹出"仍要继续吗"，用户确认才提交，取消则回到编辑器。
 * - `notify`（信息）：先提交，再弹出提示。
 */
export type EditValidationOutcome =
  | { readonly kind: "accept" }
  | {
      readonly kind: "reject" | "confirm" | "notify";
      readonly title?: string;
      readonly message: string;
    };

export const ACCEPT: EditValidationOutcome = Object.freeze({ kind: "accept" });

/** 把一次校验结果按规则的 `errorStyle` 翻成编辑器要执行的动作。 */
export function outcomeFromResult(result: DataValidationResult): EditValidationOutcome {
  if (result.valid) {
    return ACCEPT;
  }
  const kind = result.errorStyle === "warning" ? "confirm" : result.errorStyle === "information" ? "notify" : "reject";
  return {
    kind,
    ...(result.title !== undefined ? { title: result.title } : {}),
    message: result.message ?? "输入的值不符合规则"
  };
}

export interface ValidationPresentRequest {
  readonly kind: "reject" | "confirm" | "notify";
  readonly title?: string;
  readonly message: string;
  /** 被编辑单元格的矩形（编辑层坐标系）；提示贴在它下面。 */
  readonly rect: CellRect;
  /** `confirm` 专用：用户选了"继续"。 */
  readonly onContinue?: () => void;
  /** `confirm` 专用：用户选了"取消"，回到编辑器。 */
  readonly onCancel?: () => void;
}

/**
 * 提示怎么呈现不归 `EditorManager` 管——它只管什么时候要、要什么。默认实现是一个
 * DOM 气泡（`createDomValidationPresenter`），测试里换成假的就不需要 DOM。
 */
export interface ValidationPresenter {
  show(request: ValidationPresentRequest): void;
  /** 编辑单元格的位置变了（滚动、缩放），提示跟着挪。 */
  reposition(rect: CellRect): void;
  hide(): void;
  dispose(): void;
}
