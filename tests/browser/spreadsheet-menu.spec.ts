import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { WorkbookHandle } from '../../src/createWorkbook';
declare global { interface Window { menuWorkbook: WorkbookHandle } }
const item = (page: Page, id: string) => page.locator(`[data-item-id="cyexcel:${id}"]`);
async function cell(page: Page, column = 0, row = 0, right = true) {
  const box = await page.locator('#app canvas').first().boundingBox();
  if (!box) throw new Error('Missing canvas');
  await page.mouse.click(box.x + column * 110 + 30, box.y + row * 28 + 14, { button: right ? 'right' : 'left' });
}
test.beforeEach(async ({ page }) => {
  await page.goto('/tests/browser/spreadsheet-menu.html');
  await page.waitForFunction(() => Boolean(window.menuWorkbook));
});

test('default preset installs grouped flyouts, navigates by keyboard and restores focus', async ({ page }) => {
  await cell(page);
  await expect(item(page,'copy')).toBeVisible();
  await expect(item(page,'insert-row-above')).toHaveCount(0);
  await item(page,'insert').hover();
  await expect(item(page,'insert-row-above')).toBeVisible();
  await expect(page.getByRole('menu')).toHaveCount(2);
  await item(page,'insert').click();
  await expect(page.getByRole('menu').last()).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(0);
  await expect(page.locator('#app canvas').first()).toBeFocused();
  await page.keyboard.press('Shift+F10');
  await expect(item(page,'copy')).toBeVisible();
  await page.keyboard.press('End');
  await page.keyboard.press('ArrowRight');
  await expect(item(page,'sort-asc')).toBeVisible();
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(0);
});

test('clear contents and Delete preserve formatting and validation and undo restores formulas', async ({ page }) => {
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().setCell(0,0,'=2+3'));
  await cell(page);
  await item(page,'clear-contents').click();
  expect(await page.evaluate(() => {
    const sheet=window.menuWorkbook.getActiveSheet();
    const snapshot=sheet.toSnapshot();
    return {value:sheet.getCell(0,0),formula:sheet.getCellFormula(0,0),bold:sheet.getCellStyle(0,0)?.bold,
      validations:snapshot.sheets[snapshot.sheetOrder[0]!]!.dataValidations?.length};
  })).toEqual({value:undefined,formula:undefined,bold:true,validations:1});
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCellFormula(0,0))).toBe('=2+3');
  await page.keyboard.press('Delete');
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBeUndefined();
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe(5);
});

test('multi-row and multi-column insertion use the selection and undo in one step', async ({ page }) => {
  const box=await page.locator('#app canvas').first().boundingBox();
  if(!box) throw new Error('Missing canvas');
  await page.mouse.move(box.x-22,box.y+14); await page.mouse.down();
  await page.mouse.move(box.x-22,box.y+42); await page.mouse.up();
  await page.mouse.click(box.x-22,box.y+42,{button:'right'});
  await expect(item(page,'insert-row-above')).toHaveText(/插入 2 行/);
  await item(page,'insert-row-above').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(2,0))).toBe('Pear');
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Pear');
  await page.mouse.move(box.x+25,box.y-14); await page.mouse.down();
  await page.mouse.move(box.x+135,box.y-14); await page.mouse.up();
  await page.mouse.click(box.x+25,box.y-14,{button:'right'});
  await expect(item(page,'insert-column-left')).toHaveText(/插入 2 列/);
  await item(page,'insert-column-left').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,2))).toBe('Pear');
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getColumnCount())).toBe(4);
});

test('formula sheets disable structural actions with an explanation', async ({ page }) => {
  await page.goto('/tests/browser/spreadsheet-menu.html?formula');
  await page.waitForFunction(() => Boolean(window.menuWorkbook));
  await cell(page);
  await item(page,'insert').hover();
  await expect(item(page,'insert-row-above')).toHaveAttribute('aria-disabled','true');
  await expect(item(page,'insert-row-above')).toHaveAttribute('title',/公式/);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(0);
});

test('sorting moves full rows, value filter can be cleared, and clipboard failures are visible', async ({ page, context }) => {
  await cell(page,1);
  await item(page,'sort-filter').hover();
  await item(page,'sort-asc').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Apple');
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  await cell(page,2);
  await item(page,'sort-filter').hover(); await item(page,'filter-value').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().filterState !== null)).toBe(true);
  await cell(page,2);
  await item(page,'sort-filter').hover(); await item(page,'clear-filter').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().filterState)).toBeNull();
  await context.grantPermissions(['clipboard-read','clipboard-write']);
  await page.evaluate(() => navigator.clipboard.writeText('new'));
  await cell(page); await item(page,'paste').click();
  await expect.poll(() => page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('new');
  await page.evaluate(() => Object.defineProperty(navigator.clipboard,'readText',{configurable:true,value:()=>Promise.reject(new Error('Denied'))}));
  await cell(page); await item(page,'paste').click();
  await expect(page.getByRole('alert')).toContainText('Ctrl+V');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('menu')).toHaveCount(0);
});

test('small viewports scroll menus and keep flyouts within the grid', async ({ page }) => {
  await page.setViewportSize({width:400,height:360});
  await page.evaluate(() => {
    document.getElementById('app')!.style.cssText='width:350px;height:280px';
    window.menuWorkbook.getActiveSheet().resize();
  });
  await cell(page,2);
  await item(page,'insert').hover();
  await expect(page.getByRole('menu')).toHaveCount(2);
  const host=await page.locator('#app').boundingBox();
  if(!host) throw new Error('Missing host');
  for(const menu of await page.getByRole('menu').all()) {
    const box=await menu.boundingBox(); if(!box) throw new Error('Missing menu');
    expect(box.x).toBeGreaterThanOrEqual(host.x-1);
    expect(box.x+box.width).toBeLessThanOrEqual(host.x+host.width+1);
    expect(box.y+box.height).toBeLessThanOrEqual(host.y+host.height+1);
  }
  await item(page,'insert-column-right').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getColumnCount())).toBe(5);
});

test('cut waits for successful clipboard access, keeps styling, and failed paste never clears content', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read','clipboard-write']);
  await cell(page); await item(page,'cut').click();
  await expect.poll(() => page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBeUndefined();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCellStyle(0,0)?.bold)).toBe(true);
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Pear');
  await page.evaluate(() => {
    Object.defineProperty(navigator.clipboard,'writeText',{configurable:true,value:()=>Promise.reject(new Error('Denied'))});
    Object.defineProperty(navigator.clipboard,'readText',{configurable:true,value:()=>Promise.reject(new Error('Denied'))});
    Object.defineProperty(document,'execCommand',{configurable:true,value:()=>false});
  });
  await cell(page); await item(page,'cut').click();
  await expect(page.getByRole('alert')).toContainText('原数据已保留');
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Pear');
  await page.keyboard.press('Escape');
  await cell(page); await item(page,'paste').click();
  await expect(page.getByRole('alert')).toContainText('Ctrl+V');
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Pear');
});

test('merge confirmation can cancel data loss, undo restores values, and formatting is separate from content', async ({ page }) => {
  const box=await page.locator('#app canvas').first().boundingBox();
  if(!box) throw new Error('Missing canvas');
  await page.mouse.move(box.x+30,box.y+14);await page.mouse.down();
  await page.mouse.move(box.x+140,box.y+14);await page.mouse.up();
  await cell(page);
  page.once('dialog', dialog=>dialog.dismiss());
  await item(page,'merge').click();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,1))).toBe(3);
  await cell(page);
  page.once('dialog', dialog=>dialog.accept());
  await item(page,'merge').click();
  await expect.poll(() => page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,1))).toBeUndefined();
  await page.evaluate(() => window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,1))).toBe(3);
  await cell(page);
  await item(page,'clear-formats').click();
  await expect.poll(() => page.evaluate(() => window.menuWorkbook.getActiveSheet().getCellStyle(0,0)?.bold)).toBeUndefined();
  expect(await page.evaluate(() => window.menuWorkbook.getActiveSheet().getCell(0,0))).toBe('Pear');
});

test('hiding selected columns and restoring them each uses one undo step', async ({ page }) => {
  const box=await page.locator('#app canvas').first().boundingBox();
  if(!box)throw new Error('Missing canvas');
  await page.mouse.move(box.x+25,box.y-14);await page.mouse.down();
  await page.mouse.move(box.x+135,box.y-14);await page.mouse.up();
  await page.mouse.click(box.x+25,box.y-14,{button:'right'});
  await item(page,'hide-column').click();
  await expect.poll(()=>page.evaluate(()=>{
    const snapshot=window.menuWorkbook.getActiveSheet().toSnapshot();
    const columns=snapshot.sheets[snapshot.sheetOrder[0]!]!.columnData;
    return [columns[0]?.hd,columns[1]?.hd];
  })).toEqual([true,true]);
  await page.evaluate(()=>window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(()=>{
    const snapshot=window.menuWorkbook.getActiveSheet().toSnapshot();
    return Object.values(snapshot.sheets[snapshot.sheetOrder[0]!]!.columnData).some(meta=>meta.hd);
  })).toBe(false);
});

test('filtered row deletion preserves hidden records and can be undone', async ({ page }) => {
  await page.evaluate(()=>window.menuWorkbook.getActiveSheet().filterByColumn(2,value=>value==='A'));
  const box=await page.locator('#app canvas').first().boundingBox();
  if(!box)throw new Error('Missing canvas');
  await page.mouse.move(box.x-22,box.y+14);await page.mouse.down();
  await page.mouse.move(box.x-22,box.y+42);await page.mouse.up();
  await page.mouse.click(box.x-22,box.y+14,{button:'right'});
  await expect(item(page,'delete-row')).toHaveText('删除选中的可见行');
  await item(page,'delete-row').click();
  expect(await page.evaluate(()=>{
    const sheet=window.menuWorkbook.getActiveSheet();return [sheet.getCell(0,0),sheet.getCell(1,0)];
  })).toEqual(['Apple','Plum']);
  await page.evaluate(()=>window.menuWorkbook.getActiveSheet().undo());
  expect(await page.evaluate(()=>{
    const sheet=window.menuWorkbook.getActiveSheet();return [sheet.getCell(0,0),sheet.getCell(1,0),sheet.getCell(2,0)];
  })).toEqual(['Pear','Apple','Banana']);
});
