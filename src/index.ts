/**
 * @module Xlsx
 */

import XLSX , {WorkBook,Sheet, WorkSheet} from 'xlsx';
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


  @rpsAction({verbName:'read-xlsx'})
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





  private getSheet(ctx:RpsContext, sheetNameNum:string|number) : Sheet{
    let wb:WorkBook = ctx.getModuleContext(MOD_ID)['workbook'];
    let sheet = typeof sheetNameNum === "string" ? 
      wb.Sheets[sheetNameNum] : R.values(wb.Sheets)[sheetNameNum-1];
    return sheet
  }


}
