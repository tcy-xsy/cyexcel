import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { ModelRevision } from "../model";
import type { MutationPlanningResources } from "../mutations/types";
import type { SheetModel } from "./SheetModel";
import { SheetTransaction } from "./SheetTransaction";

/**
 * 一次事务规划所需的全部资源：一个影子模型，加上写口与两个读口。
 *
 * 影子在这里克隆出来（`SheetTransaction` 的构造函数），此后到 commit 之前活模型
 * 一个字节都不会动。`baseRevision` 直接交给它，所以不存在「事务的基准版本与协调器
 * 对不上」这种情况——那是数据源自己开事务时代的检查。
 *
 * 三个口共用同一条**归属校验**：一个协调器一次只服务一张表，交进来的模型不是那一
 * 张就说明调用方拿错了对象。读也要查——拿错模型的调用方读到的是**另一张表**的行高，
 * 而它接着会把这个数当成「这一行原来多高」写进逆里，比写错一张表更难查。
 */
export function createSheetPlanningResources(
  sheet: SheetModel,
  baseRevision: ModelRevision
): MutationPlanningResources {
  const transaction = new SheetTransaction(sheet, baseRevision);
  const assertOwned = (source: SheetModel): void => {
    if (source !== sheet) {
      throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Mutation targets another sheet");
    }
  };
  return {
    model: transaction,
    prepareData: (source, mutation) => {
      assertOwned(source);
      return transaction.prepareMutation(mutation);
    },
    readRecord: (source, row, column) => {
      assertOwned(source);
      return transaction.getRecord(row, column);
    },
    readRowMeta: (source, row) => {
      assertOwned(source);
      return transaction.getRowMeta(row);
    },
    readColumnMeta: (source, column) => {
      assertOwned(source);
      return transaction.getColumnMeta(column);
    },
    dispose: () => {
      transaction.dispose();
    }
  };
}
