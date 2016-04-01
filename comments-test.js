"use strict";
const fs = require('fs');
const XLSX = require('./xlsx');
const expect = require('chai').expect;

function writeAndReadContent(wb, fname) {
  XLSX.writeFile(wb, fname);
  return XLSX.readFile(fname);
}

describe('XLSX comments', () => {
  context('Duplicate a xlsx file with comments', () => {

    let output;

    beforeEach(() => {
      output = '/tmp/' + new Date().getTime() + '.xlsx';
    });

    afterEach(() => {
      //fs.unlinkSync(output);
    });

    it('should have the same content as the original', () => {
      const fname = './comments-test.xlsx';
      const input_wb = XLSX.readFileSync(fname);
      const output_wb = writeAndReadContent(input_wb, output);

      expect(output_wb.Directory.comments)
        .to.deep.equal(input_wb.Directory.comments);

      // ignoring author ('a')
      const keys = ['h', 'r', 't'];
      keys.forEach(key => {
        expect(output_wb.Sheets.Sheet1.B2.c[key])
        .to.equal(input_wb.Sheets.Sheet1.B2.c[key]);
      })
    });
  });
});
