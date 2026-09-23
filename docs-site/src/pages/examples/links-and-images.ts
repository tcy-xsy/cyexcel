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
  { width: 72, header: "图片", sortable: false, filterable: false, renderer: createImageCellRenderer() }
];

const rows = PRODUCTS.map((name, index) => [
  name,
  `https://www.cyexcel.example/products/CY-${String(index + 1).padStart(6, "0")}`,
  productIconDataUri(index)
]);

export default function linksAndImagesExamplePage(root: HTMLElement): () => void {
  root.append(
    pageHeader(
      "EXAMPLES",
      "链接与图片单元格",
      "两条刻意不对称的机制：链接是看起来像就自动生效，图片必须显式声明。"
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

  const imageSection = section("图片单元格");
  imageSection.body.append(
    paragraph(
      "图片则相反：必须显式把 createImageCellRenderer() 配到列的 renderer 上，那一列的值（" +
        "http(s) URL 或 data: URI）才会被当图片画出来，等比缩放、居中，不会裁切。不会因为文本 " +
        "“看起来像图片”（比如以 .png 结尾）就自动生效——那样会和普通链接文本冲突。"
    )
  );
  root.append(imageSection.root);
  root.append(codeBlock(`import { createImageCellRenderer } from "@cyexcel/core";

const columns = [
  { width: 72, header: "图片", renderer: createImageCellRenderer() }
  // ...
];
sheet.setCell(row, column, "https://example.com/avatar.png");`));

  const demo = demoBlock({
    title: "两列的实际效果",
    description: "“官网”列的每一格都是纯文本 URL；“图片”列配了 createImageCellRenderer()。"
  });
  root.append(demo.root);

  const sheet = createSheet(demo.mount, {
    columns,
    snapshot: matrixToWorkbookSnapshot(rows),
    headerHeight: 28,
    rowHeaderWidth: 48,
    rowHeight: 40
  });

  return () => sheet.dispose();
}
