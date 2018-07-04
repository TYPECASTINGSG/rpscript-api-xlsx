import {expect} from 'chai';
import m from 'mocha';

import RPSXlsx from '../src/index';
import { RpsContext } from 'rpscript-interface';

m.describe('XLSX', () => {

  m.it('should xlsx', async function () {
    let xlsx = new RPSXlsx;

    let output = await xlsx.readXlsx(new RpsContext,{},"formula_stress_test.xlsx");
    
    console.log(output.SheetNames);

  }).timeout(0);

})
