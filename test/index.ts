import {expect} from 'chai';
import m from 'mocha';
import R from 'ramda';
import RPSXlsx from '../src/index';
import { RpsContext } from 'rpscript-interface';

m.describe('XLSX', () => {

  m.it('should xlsx', async function () {
    let ctx = new RpsContext
    let xlsx = new RPSXlsx(ctx);

    let output = await xlsx.readXlsx(ctx,{},"formula_stress_test.xlsx");
    
    console.log(output.SheetNames);
    // let sheet = output.Sheets['Logical'];
    let html:any = await xlsx.sheetToCsv(ctx,{},'Logical');
    
    html = await xlsx.sheetToFormulae(ctx,{},'Logical');

    html = await xlsx.sheetToTxt(ctx,{},'Logical');

  }).timeout(0);

})
