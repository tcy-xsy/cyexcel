import { createSheet } from "@cyexcel/createSheet";
import { createImageCellRenderer, matrixToWorkbookSnapshot } from "@cyexcel/core";
import { codeBlock, demoBlock, pageHeader, paragraph, section } from "../../components/doc";

const PRODUCTS = ["工业温度传感器", "户外电源模组", "智能门锁套件", "车载充电底座", "空气质量探头"];
const AVATAR_COLORS = ["#f97316", "#3b82f6", "#10b981", "#a855f7", "#ef4444"];

/** 纯本地 data: URI 小图标，文档站不依赖任何外部图片服务就能演示图片单元格。 */
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
  // undefined，不是 ""——这一格在导入时压根不占用（matrixToWorkbookSnapshot 的
  // 口径：null/undefined 是"没有这一格"，"" 是一个真实的空字符串值）。图片单元格
  // 就该是"没有 v，只有 img"，跟 Univer 的 insertCellImageByUrl() 只设 p/t、
  // 从不设 v 是同一回事，见下面 setCellImage() 的调用。
  undefined
]);

export default function linksAndImagesExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "链接与图片单元格",
      "三条刻意不对称的机制：链接是看起来像就自动生效，图片（列配置）必须显式声明整列，图片（单元格）则是某一格自己决定要不要显示。"
    )
  );

  const linkSection = section("链接自动识别");
  linkSection.body.append(
    paragraph(
      "不需要给列或单元格做任何配置。只要一个单元格的完整内容（去掉首尾空白）本身就是一个 " +
        "URL 或邮箱地址，默认绘制路径就会自动画成蓝色下划线，点击会在新标签页打开——这是 Excel " +
        "“AutoFormat As You Type” 的规则：整格都是链接才生效，句子里嵌一个 URL 不会被识别。" +
        "不会改动单元格的原始值，纯粹是渲染 + 点击这层的行为。"
    )
  );
  root.append(linkSection.root);
  root.append(codeBlock(`sheet.setCell(row, column, "https://www.cyexcel.example/products/CY-000001");
// 不用配 renderer，也不用声明列类型——整格是 URL 就自动画成链接、可点击打开。`));

  const imageSection = section("图片单元格（列配置）");
  imageSection.body.append(
    paragraph(
      "图片则相反：必须显式把 createImageCellRenderer() 配到列的 renderer 上，那一列的值（" +
        "http(s) URL 或 data: URI）才会被当图片画出来，等比缩放、居中，不会裁切。不会因为文本 " +
        "“看起来像图片”（比如以 .png 结尾）就自动生效——那样会和普通链接文本冲突。" +
        "这条路径适合“整列本来就是图片字段”的批量场景（比如从业务系统导入的商品图库列）。"
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
      "Excel、Univer 的图片单元格是这一格自己的属性，不是列的属性——同一列里，有的格是文字、" +
        "有的格是图片，互不干扰。cyexcel 用 setCellImage()/getCellImage() 表达同样的心智模型：" +
        "列上什么 renderer 都不用配，直接对某一格调用 setCellImage() 就够了，图片会优先于该列" +
        "其他任何渲染方式画出来。清空图片传 undefined 即可，不影响这一格的值（v）与样式（s）。"
    )
  );
  root.append(cellImageSection.root);
  root.append(codeBlock(`const columns = [
  { width: 72, header: "图片" }
  // 不需要配 renderer——这一列其余格照常显示文字
];

// 只有这一格显示图片，同一列的其他格不受影响
sheet.setCellImage(row, column, "https://example.com/avatar.png");
// 清空
sheet.setCellImage(row, column, undefined);`));

  const demo = demoBlock({
    title: "三列的实际效果",
    description:
      "“官网”列是纯文本 URL；“图片（列配置）”列配了 createImageCellRenderer()；" +
      "“图片（单元格）”列什么 renderer 都没配，只有第 1、3 行用 setCellImage() 单独设了图片。"
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
