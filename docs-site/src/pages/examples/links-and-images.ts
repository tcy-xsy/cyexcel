import { createSheet } from "@cyexcel/createSheet";
import { createImageCellRenderer, matrixToWorkbookSnapshot } from "@cyexcel/core";
import { codeBlock, demoBlock, pageHeader, paragraph, section } from "../../components/doc";

const PRODUCTS = ["工业温度传感器", "户外电源模组", "智能门锁套件", "车载充电底座", "空气质量探头"];
const AVATAR_COLORS = ["#f97316", "#3b82f6", "#10b981", "#a855f7", "#ef4444"];

/** 使用本地 data: URI 图标，避免示例依赖外部图片服务。 */
function productIconDataUri(index: number): string {
  const color = AVATAR_COLORS[index % AVATAR_COLORS.length];
  const letter = PRODUCTS[index % PRODUCTS.length]!.charAt(0);
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64">` +
    `<rect width="64" height="64" rx="12" fill="${color}"/>` +
    `<text x="32" y="42" font-size="26" text-anchor="middle" fill="#fff" font-family="sans-serif">${letter}</text>` +
    `</svg>`;
  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}

const columns = [
  { width: 200, header: "产品", sortable: false, filterable: false },
  { width: 260, header: "官网", sortable: false, filterable: false },
  { width: 72, header: "图片（列配置）", sortable: false, filterable: false, renderer: createImageCellRenderer() },
  { width: 72, header: "图片（单元格）", sortable: false, filterable: false }
];

const rows = PRODUCTS.map((name, index) => [
  name,
  `https://www.cyexcel.example/products/CY-${String(index + 1).padStart(6, "0")}`,
  productIconDataUri(index),
  // undefined 不生成值记录，图片在创建表格后通过 setCellImage() 设置。
  undefined
]);

export default function linksAndImagesExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "链接与图片单元格",
      "支持链接自动识别、列级图片渲染和单元格图片设置。"
    )
  );

  const linkSection = section("链接自动识别");
  linkSection.body.append(
    paragraph(
      "单元格内容去除首尾空白后为完整 URL 或邮箱地址时，默认以蓝色下划线显示，点击可在新标签页打开。" +
        "文本中嵌入的 URL 不会被自动识别。链接识别不修改单元格原始值，无需额外配置列类型或渲染器。"
    )
  );
  root.append(linkSection.root);
  root.append(codeBlock(`sheet.setCell(row, column, "https://www.cyexcel.example/products/CY-000001");
// 完整 URL 自动识别为可点击链接。`));

  const imageSection = section("图片单元格（列配置）");
  imageSection.body.append(
    paragraph(
      "为列的 renderer 配置 createImageCellRenderer() 后，该列中的 HTTP(S) URL 或 data: URI 将作为图片显示。" +
        "图片保持宽高比并居中显示，不进行裁切。此方式适用于整列存储图片地址的场景。普通文本不会仅因包含图片扩展名而显示为图片。"
    )
  );
  root.append(imageSection.root);
  root.append(codeBlock(`import { createImageCellRenderer } from "@cyexcel/core";

const columns = [
  { width: 72, header: "图片", renderer: createImageCellRenderer() }
  // ...
];
sheet.setCell(row, column, "https://example.com/avatar.png");`));

  const cellImageSection = section("图片单元格（单元格级）");
  cellImageSection.body.append(
    paragraph(
      "setCellImage() 和 getCellImage() 用于设置和读取单元格图片，无需配置列渲染器。" +
        "同一列中的文字和图片可分别设置，单元格图片优先于列渲染器显示。" +
        "传入 undefined 可清除图片，不影响单元格的值（v）与样式（s）。"
    )
  );
  root.append(cellImageSection.root);
  root.append(codeBlock(`const columns = [
  { width: 72, header: "图片" }
  // 无需配置 renderer，同列其他单元格仍按默认方式渲染。
];

// 仅为指定单元格设置图片。
sheet.setCellImage(row, column, "https://example.com/avatar.png");
// 清空
sheet.setCellImage(row, column, undefined);`));

  const demo = demoBlock({
    title: "链接与图片示例",
    description:
      "“官网”列演示 URL 自动识别；“图片（列配置）”列使用 createImageCellRenderer()；" +
      "“图片（单元格）”列通过 setCellImage() 为第 1、3 行设置图片。"
  });
  root.append(demo.root);

  const sheet = createSheet(demo.mount, {
    columns,
    snapshot: matrixToWorkbookSnapshot(rows),
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 40
  });
  sheet.setCellImage(0, 3, productIconDataUri(0));
  sheet.setCellImage(2, 3, productIconDataUri(2));

  return () => sheet.dispose();
}
