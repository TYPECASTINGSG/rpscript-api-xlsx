import {expect} from 'chai';
import m from 'mocha';
import R from 'ramda';
import RPSXlsx from '../src/index';
import { RpsContext } from 'rpscript-interface';

m.describe('XLSX', () => {

  m.it('should xlsx', async function () {
    let ctx = new RpsContext
    let xlsx = new RPSXlsx(ctx);

    let output = await xlsx.readXlsx(ctx,{},"TOBEDELETED.xlsx");
    
    console.log(output.SheetNames);

    let sheet = await xlsx.readSheet(ctx,{},"Lapsed Policies");
    console.log('old range : '+sheet['!ref']);
    console.log('cols : '+sheet['!cols']);
    console.log('rows : '+sheet['!rows']);
    console.log('merges : '+sheet['!merges']);
    console.log('protect : '+sheet['!protect']);
    console.log('autofilter : '+sheet['!autofilter']);

    //  await xlsx.addColumn(ctx,{},"Lapsed Policies","ABC",1);
    //  await xlsx.addRow(ctx,{},"Lapsed Policies",new Date);
    //  await xlsx.setCell(ctx,{},"Lapsed Policies","T4","HEllo");

     console.log('new range : '+sheet['!ref']);
    sheet = await xlsx.sheetToFormulae(ctx,{},"Lapsed Policies");
    let cell = xlsx.getCells(ctx,{column:'T'},"Lapsed Policies");
    // let result = await xlsx.sheetToCsv(ctx,{},"Lapsed Policies");

    console.log(cell);
    // let t = ctx.getModuleContext('xlsx')['workbook'].Sheets['Lapsed Policies'];

    // await xlsx.writeXlsx(ctx,{},'TOBEDELETED2.xlsx');

  }).timeout(0);

})
