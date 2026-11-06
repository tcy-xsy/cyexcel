import { describe, expect, it, vi } from 'vitest';
import { createBasicContextMenuPlugin } from '../../../../src/core/menu/basicMenuPlugin';
import type { ContextMenuContext, ContextMenuItem, ContextMenuProvider, SheetMenuCommands } from '../../../../src/core/menu/types';
import type { PluginContext } from '../../../../src/core/plugins';

function setup() {
  let provider: ContextMenuProvider | undefined;
  const run = vi.fn();
  const state = {structureEnabled:true,canMerge:true,hasMerge:false,canSort:true,canFilter:true,hasFilter:false,hasHiddenColumns:false,bold:false,italic:false};
  const context = {
    contextMenu:{registerProvider:(_meta:unknown, value:ContextMenuProvider)=>{provider=value;return ()=>{};}},
    sheetMenu: {getState:()=>state,run} satisfies SheetMenuCommands,
    copy:vi.fn(async()=>({status:'copied'})),paste:vi.fn(async()=>{}),
    insertRows:vi.fn(),deleteRows:vi.fn(),insertColumns:vi.fn(),deleteColumns:vi.fn(),hideColumn:vi.fn()
  };
  createBasicContextMenuPlugin().setup(context as unknown as PluginContext);
  const range={startRow:2,endRow:4,startColumn:1,endColumn:2};
  const selection={kind:'cell' as const,active:{row:2,column:1},anchor:{row:2,column:1},range,ranges:[range]};
  const menu:ContextMenuContext={target:'cell',row:2,column:1,selection,point:{x:0,y:0}};
  const find=(id:string,items:readonly ContextMenuItem[]=provider!(menu)):ContextMenuItem|undefined=>{
    for(const item of items){if(item.id===`cyexcel:${id}`)return item;if(item.type==='submenu'){const child=find(id,item.children);if(child)return child;}}
    return undefined;
  };
  return {context,state,menu,find,provider:provider!};
}

describe('spreadsheet menu commands',()=>{
  it('inserts and deletes the full continuous selection rather than just the clicked row',()=>{
    const {context,find,menu}=setup();
    const signal=new AbortController().signal;
    for(const id of ['insert-row-above','insert-column-right','delete-row']){
      const item=find(id);if(item?.type!=='action')throw new Error('Missing action');item.run(menu,signal);
    }
    expect(context.insertRows).toHaveBeenCalledWith(2,3);
    expect(context.insertColumns).toHaveBeenCalledWith(3,[{},{}]);
    expect(context.deleteRows).toHaveBeenCalledWith(2,3);
  });
  it('passes the opening selection and cancellation signal to async clipboard commands',async()=>{
    const {context,find,menu}=setup();
    const signal=new AbortController().signal;
    const item=find('copy');if(item?.type!=='action')throw new Error('Missing action');
    await item.run(menu,signal);
    expect(context.copy).toHaveBeenCalledWith({range:menu.selection.range,signal});
  });
  it('disables reference-changing commands on formula sheets and rejects partial selections explicitly',()=>{
    const {state,find,menu,provider}=setup();state.structureEnabled=false;
    expect(find('insert-row-above')).toMatchObject({disabled:true,description:expect.stringContaining('公式')});
    expect(find('sort-asc')).toMatchObject({disabled:true});
    const items=provider({...menu,selection:{...menu.selection,ranges:[menu.selection.range,menu.selection.range]}});
    expect(find('copy',items)).toMatchObject({disabled:true});
    expect(find('clear-contents',items)).toMatchObject({disabled:true});
  });
  it('omits spreadsheet operations outside the grid',()=>{
    const {menu,provider}=setup();
    expect(provider({...menu,target:'blank'})).toEqual([]);
  });
});
