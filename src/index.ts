/**
 * @module Xlsx
 */

import XLSX , {WorkBook} from 'xlsx';
import {RpsContext,RpsModule,rpsAction} from 'rpscript-interface';

@RpsModule("xlsx")
export default class RPSXlsx {

  @rpsAction({verbName:'readXLSX'})
  async  readXlsx(ctx:RpsContext,opts:Object, filename:any) : Promise<WorkBook>{
    return XLSX.readFile(filename);
  }

}
