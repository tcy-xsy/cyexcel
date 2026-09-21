/**
 * 工具栏图标：内联 SVG 串。
 *
 * 用 stroke 风格、24×24、`currentColor`，与项目现有界面外框的"细而轻"一致
 * （见 render/theme.ts 的排序指示器注释）。内联而不是引图标库：一个图标库的体积
 * 顶得上整个工具栏模块，而我们只要这二十来个。
 *
 * 它们只会被 `innerHTML` 到我们自己创建的元素里，内容全是字面量、没有任何外部输入
 * 参与拼接——这是这里能用 innerHTML 的前提。
 */

const wrap = (body: string): string =>
  `<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;

export const undo = wrap('<path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13"/>');
export const redo = wrap('<path d="M21 7v6h-6"/><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3L21 13"/>');
export const copy = wrap('<rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/>');
export const paste = wrap('<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>');
export const clearStyle = wrap('<path d="M4 20h16"/><path d="M6 16 16 6a2.8 2.8 0 0 0-4-4L2 12"/><path d="m10 6 4 4"/>');

export const bold = wrap('<path d="M7 5h6a3.5 3.5 0 0 1 0 7H7z"/><path d="M7 12h7a3.5 3.5 0 0 1 0 7H7z"/>');
export const italic = wrap('<path d="M15 4h-6"/><path d="M13 20H7"/><path d="M14 4 10 20"/>');
export const underline = wrap('<path d="M7 4v6a5 5 0 0 0 10 0V4"/><path d="M5 20h14"/>');
export const strike = wrap('<path d="M5 12h14"/><path d="M8 7a4 4 0 0 1 8 0"/><path d="M8 16a4 4 0 0 0 8 0"/>');
export const textColor = wrap('<path d="m5 17 5-12 5 12"/><path d="M7 13h6"/><path d="M4 21h16"/>');
export const background = wrap('<path d="M12 3 4 11l6 6 8-8z"/><path d="M6 21h14"/>');
export const border = wrap('<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 12h18"/><path d="M12 3v18"/>');

export const percent = wrap('<circle cx="7" cy="7" r="2.5"/><circle cx="17" cy="17" r="2.5"/><path d="M19 5 5 19"/>');
export const thousands = wrap('<path d="M4 7h16"/><path d="M4 12h10"/><path d="M4 17h16"/><circle cx="18" cy="12" r="1"/>');
export const addDecimalIcon = wrap('<path d="M5 18h1"/><path d="M9 6v12"/><path d="M13 6v12"/><path d="M17 9v6"/><path d="M20 12h-6"/>');
export const subtractDecimalIcon = wrap('<path d="M5 18h1"/><path d="M9 6v12"/><path d="M13 6v12"/><path d="M20 12h-6"/>');

export const alignLeft = wrap('<path d="M4 6h16"/><path d="M4 12h10"/><path d="M4 18h13"/>');
export const alignCenter = wrap('<path d="M4 6h16"/><path d="M7 12h10"/><path d="M6 18h12"/>');
export const alignRight = wrap('<path d="M4 6h16"/><path d="M10 12h10"/><path d="M7 18h13"/>');
export const alignTop = wrap('<path d="M4 4h16"/><path d="M12 8v12"/><path d="m8 12 4-4 4 4"/>');
export const alignMiddle = wrap('<path d="M4 12h16"/><path d="M12 4v4"/><path d="M12 16v4"/>');
export const alignBottom = wrap('<path d="M4 20h16"/><path d="M12 4v12"/><path d="m8 12 4 4 4-4"/>');

export const sortAsc = wrap('<path d="M7 20V4"/><path d="m3 8 4-4 4 4"/><path d="M14 7h7"/><path d="M14 12h5"/><path d="M14 17h3"/>');
export const sortDesc = wrap('<path d="M7 4v16"/><path d="m3 16 4 4 4-4"/><path d="M14 7h3"/><path d="M14 12h5"/><path d="M14 17h7"/>');
export const filter = wrap('<path d="M3 5h18l-7 8v6l-4 2v-8z"/>');
export const find = wrap('<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>');

export const merge = wrap('<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M9 5v14"/><path d="M15 5v14"/><path d="M7 12h10"/>');
export const unmerge = wrap('<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M12 5v14"/>');
export const insertRow = wrap('<rect x="3" y="12" width="18" height="8" rx="1"/><path d="M12 4v5"/><path d="M9.5 6.5h5"/>');
export const deleteRow = wrap('<rect x="3" y="12" width="18" height="8" rx="1"/><path d="M9.5 7h5"/>');
export const insertColumn = wrap('<rect x="12" y="3" width="8" height="18" rx="1"/><path d="M5 12h5"/><path d="M7.5 9.5v5"/>');
export const deleteColumn = wrap('<rect x="12" y="3" width="8" height="18" rx="1"/><path d="M5 12h5"/>');
export const hideColumn = wrap('<path d="M3 3l18 18"/><path d="M10.6 5.1A9 9 0 0 1 21 12a9 9 0 0 1-3 3.9"/><path d="M6.6 6.6A9 9 0 0 0 3 12a9 9 0 0 0 9 6 9 9 0 0 0 3.2-.6"/>');
export const caretDown = wrap('<path d="m7 10 5 5 5-5"/>');
export const more = wrap('<circle cx="5" cy="12" r="1.5"/><circle cx="12" cy="12" r="1.5"/><circle cx="19" cy="12" r="1.5"/>');
