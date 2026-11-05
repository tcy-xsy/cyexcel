import type { DataValidationResult } from "../sheet/validations";

/**
 * 编辑器提交前的校验结论（数据验证设计稿 §7.1）。`EditorManager` 只认这个形状，
 * 不认识规则本身：规则怎么查、值怎么判是集成层（`CyExcelCore`）的事。
 *
 * 与 Univer 一致，只有**停止**级（`errorStyle: "stop"`，默认）会拦：
 * - `accept`：照常提交。违规但不是停止级的值也走这里——只会被标上红色小三角。
 * - `reject`：不提交，弹出模态对话框说明原因；关闭时放弃本次修改，保留原值并退出编辑。
 *
 * 警告、信息两档不弹确认、也不提示：Univer 的拦截器只处理 `STOP`，其余直接放行。
 */
export type EditValidationOutcome =
  | { readonly kind: "accept" }
  | {
      readonly kind: "reject";
      readonly title?: string;
      readonly message: string;
    };

export const ACCEPT: EditValidationOutcome = Object.freeze({ kind: "accept" });

/** 把一次校验结果按规则的 `errorStyle` 翻成编辑器要执行的动作：只有停止级拦。 */
export function outcomeFromResult(result: DataValidationResult): EditValidationOutcome {
  if (result.valid || (result.errorStyle ?? "stop") !== "stop") {
    return ACCEPT;
  }
  return {
    kind: "reject",
    ...(result.title !== undefined ? { title: result.title } : {}),
    message: result.message ?? "输入的值不符合规则"
  };
}

export interface ValidationPresentRequest {
  readonly title?: string;
  readonly message: string;
  /** 用户点了「确定」（或按了 Esc）。 */
  readonly onClose: () => void;
}

/**
 * 拒绝输入时怎么呈现不归 `EditorManager` 管——它只管什么时候要、要什么。默认实现是一个
 * 模态对话框（`createDomValidationPresenter`），测试里换成假的就不需要 DOM。
 */
export interface ValidationPresenter {
  show(request: ValidationPresentRequest): void;
  hide(): void;
  dispose(): void;
}
