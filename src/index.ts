/**
 * @module Xlsx
 */

import XLSX , {WorkBook,Sheet, WorkSheet,Range} from 'xlsx';
import R from 'ramda';
import {RpsContext,RpsModule,rpsAction} from 'rpscript-interface';

let MOD_ID = 'xlsx'

export interface XlsxContext {
  workbook?:WorkBook
}

@RpsModule(MOD_ID)
export default class RPSXlsx {

  constructor(ctx:RpsContext){
    ctx.addModuleContext(MOD_ID,{});
  }


  @rpsAction({verbName:'read-workbook'})
  async  readXlsx(ctx:RpsContext,opts:Object, filename:any) : Promise<WorkBook>{
    let workbook = XLSX.readFile(filename);

    ctx.addModuleContext(MOD_ID,{workbook:workbook});

    return workbook;
  }

  @rpsAction({verbName:'write-xlsx'})
  async  writeXlsx(ctx:RpsContext,opts:Object, filename:any) : Promise<WorkBook>{
    let wb = ctx.getModuleContext(MOD_ID)['workbook'];
    XLSX.writeFile(wb,filename);

    ctx.addModuleContext(MOD_ID,{});

    return wb;
  }

  @rpsAction({verbName:'create-new-xlsx'})
  async  newXlsx(ctx:RpsContext,opts:Object) : Promise<WorkBook>{
    let workbook = XLSX.utils.book_new();
    ctx.addModuleContext(MOD_ID,{workbook:workbook});

    return workbook;
  }

  @rpsAction({verbName:'read-sheet'})
  async  readSheet(ctx:RpsContext,opts:Object, sheetname:string|number) : Promise<Sheet>{
    return this.getSheet(ctx,sheetname);
  }

  @rpsAction({verbName:'copy-sheet-from-workbook'})
  async  copySheet(ctx:RpsContext,opts:Object, workbook:string, sheetname:string|number, toSheetname:string|number) : Promise<Sheet>{
    return this.getSheet(ctx,sheetname);
  }

  //convert sheet to others (exporting)

  @rpsAction({verbName:'convert-sheet-to-csv'})
  async  sheetToCsv(ctx:RpsContext,opts:Object, sheetNameNum:string|number) : Promise<string>{
    return XLSX.utils.sheet_to_csv( this.getSheet(ctx,sheetNameNum) );
  }
  @rpsAction({verbName:'convert-sheet-to-text'})
  async  sheetToTxt(ctx:RpsContext,opts:Object, sheetNameNum:string) : Promise<string>{
    return XLSX.utils.sheet_to_txt( this.getSheet(ctx,sheetNameNum) );
  }
  @rpsAction({verbName:'convert-sheet-to-html'})
  async  sheetToHtml(ctx:RpsContext,opts:Object, sheetNameNum:string) : Promise<string>{
    return XLSX.utils.sheet_to_html( this.getSheet(ctx,sheetNameNum) );
  }
  @rpsAction({verbName:'convert-sheet-to-json'})
  async  sheetToJson(ctx:RpsContext,opts:Object, sheetNameNum:string) : Promise<{}[]>{
    return XLSX.utils.sheet_to_json( this.getSheet(ctx,sheetNameNum) );
  }
  @rpsAction({verbName:'convert-sheet-to-formulae'})
  async  sheetToFormulae(ctx:RpsContext,opts:Object, sheetNameNum:string) : Promise<string[]>{
    return XLSX.utils.sheet_to_formulae( this.getSheet(ctx,sheetNameNum) );
  }

  //column name, 
  @rpsAction({verbName:'add-xlsx-column'})
  async  addColumn(ctx:RpsContext,opts:Object, 
    sheetNameNum:string, colName:string, val:string|number|Date|Function|Array<any> ) : Promise<any>{
    let sheet = this.getSheet(ctx,sheetNameNum);
    let range = this.getRange(ctx,sheetNameNum);
    let newCol= range.e.c + 1;

    let newStartRange = XLSX.utils.encode_cell({r: range.s.r, c: range.s.c});
    let newRange = XLSX.utils.encode_cell({r: range.e.r, c: newCol});
    sheet['!ref'] = newStartRange+':'+newRange;

    for(let r=range.s.r;r < range.e.r; r++){
      let a1 = XLSX.utils.encode_cell({r: r, c: newCol})

      if(r===range.s.r) sheet[a1] = {v: colName, t: 's'};
      else sheet[a1] = this.encodeCellValue(val);
    }
  }
  @rpsAction({verbName:'add-xlsx-row'})
  async addRow(ctx:RpsContext,opts:Object, sheetNameNum:string, 
    val:string|number|Date|Function|Array<any>) : Promise<any>{
      let sheet = this.getSheet(ctx,sheetNameNum);
      let range = this.getRange(ctx,sheetNameNum);
      let newRow= range.e.r + 1;
  
      let newStartRange = XLSX.utils.encode_cell({r: range.s.r, c: range.s.c});
      let newRange = XLSX.utils.encode_cell({r: newRow, c: range.e.c});
      sheet['!ref'] = newStartRange+':'+newRange;
  
      for(let c=range.s.c;c < range.e.c; c++){
        let a1 = XLSX.utils.encode_cell({r: newRow, c: c});
        sheet[a1] = this.encodeCellValue(val);
      }
  }
  @rpsAction({verbName:'add-xlsx-cell'})
  async  addCell(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{
  }

  encodeCellValue (value:string|number|Function|Date|Array<any>) : Object{
    let encoding;
    if(typeof value === 'number') encoding = {v:value,t:'n'};
    else if(typeof value === 'string') encoding = {v:value,t:'s'};
    else if(value instanceof Date) encoding = {v:value,t:'d'};

    return encoding;
  }

  @rpsAction({verbName:'set-xlsx-column'})
  async  setColumn(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{}
  @rpsAction({verbName:'set-xlsx-row'})
  async  setRow(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{}
  @rpsAction({verbName:'set-xlsx-cell'})
  async  setCell(ctx:RpsContext,opts:Object, sheetNameNum:string, col:string,
  value:string|number|Function) : Promise<any>{
    let sheet = this.getSheet(ctx,sheetNameNum);
    // let a1 = XLSX.utils.encode_cell({r: newRow, c: c});
    sheet[col] = this.encodeCellValue(value);

  }

  @rpsAction({verbName:'remove-xlsx-column'})
  async  deleteColumn(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{}
  @rpsAction({verbName:'remove-xlsx-row'})
  async  deleteRow(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{}
  @rpsAction({verbName:'remove-xlsx-cell'})
  async  deleteCell(ctx:RpsContext,opts:Object, sheetNameNum:string, colName:string) : Promise<any>{}

  //convert others to sheet (importing)

  @rpsAction({verbName:'convert-aoa-to-sheet'})
  async  aoaToSheet(ctx:RpsContext,opts:Object, data:any[][]) : Promise<WorkSheet>{
    return XLSX.utils.aoa_to_sheet(data);
  }
  @rpsAction({verbName:'convert-json-to-sheet'})
  async  jsonToSheet(ctx:RpsContext,opts:Object, json:{}[]) : Promise<WorkSheet>{
    return XLSX.utils.json_to_sheet(json);
  }
  @rpsAction({verbName:'convert-sheet-add-aoa'})
  async  sheetAddAoa(ctx:RpsContext,opts:Object, sheetNameNum:string|number,data:any[][]) : Promise<WorkSheet>{
    let sheet = XLSX.utils.sheet_to_formulae( this.getSheet(ctx,sheetNameNum) );
    let addedSheet = XLSX.utils.sheet_add_aoa(sheet,data);

    return addedSheet;
  }
  @rpsAction({verbName:'convert-sheet-add-json'})
  async  sheetAddJson(ctx:RpsContext,opts:Object, sheetNameNum:string|number, json:{}[]) : Promise<WorkSheet>{
    let sheet = XLSX.utils.sheet_to_formulae( this.getSheet(ctx,sheetNameNum) );
    let addedSheet = XLSX.utils.sheet_add_json(sheet,json);
    
    return addedSheet;
  }

  // cell manipulation

  // @rpsAction({verbName:'format-cell'})
  // async formatCell(ctx:RpsContext,opts:Object, sheetNameNum:string|number, json:{}[]) : Promise<WorkSheet>{
  //   let sheet = XLSX.utils.sheet_to_formulae( this.getSheet(ctx,sheetNameNum) );
  //   let addedSheet = XLSX.utils.
  // }


  getCells (ctx:RpsContext,opt:Object,sheetNameNum:string|number, a1?:string) : Object|Array<Object>{
    let sheet = this.getSheet(ctx,sheetNameNum);
    if(opt['column']){
      let matchColumn = (val, key) => opt['column'] == key.charAt(0);

      return R.pickBy(matchColumn, sheet);
    }else if(opt['row']){
      let matchColumn = (val, key) => { 
        let v = key.split(':');
        return opt['row'] == v[1];
      };

      return R.pickBy(matchColumn, sheet);
    }
    else if(opt['header']){
    }
    
    return sheet[a1];
  }
  getCol (ctx:RpsContext,sheetNameNum:string|number) : number{
    let sheet = this.getSheet(ctx,sheetNameNum);
    return XLSX.utils.decode_col(sheet['!ref']);
  }
  getRange (ctx:RpsContext,sheetNameNum:string|number) : Range{
    let sheet = this.getSheet(ctx,sheetNameNum);
    return XLSX.utils.decode_range(sheet['!ref']);
  }

  private getSheet(ctx:RpsContext, sheetNameNum:string|number) : Sheet{
    let wb:WorkBook = ctx.getModuleContext(MOD_ID)['workbook'];
    let sheet = typeof sheetNameNum === "string" ? 
      wb.Sheets[sheetNameNum] : R.values(wb.Sheets)[sheetNameNum-1];
    return sheet
  }

  // private getRowByHeaderName (name:string) {
    // let sheet = this.getSheet(ctx,sheetNameNum);
    // let firstRow = sheet["!ref"].split(':').filter('not digit');
    // let firstRow = 1;
  // }


}
