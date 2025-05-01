async function main(workbook: ExcelScript.Workbook, startRow: number) {
    try {
      const tableOld = workbook.getTable("tableOld");
      const desiredEndRow = 1000;
      if (!tableOld) {
        throw new Error("Table 'tableOld' not found");
      }
  
      // テーブルのデータを一度に取得
      //const oldTableData = tableOld.getRangeBetweenHeaderAndTotal().getValues();
      const oldTableData = tableOld.getRange().getValues();
      const actualOldRowCount = oldTableData.length;
      console.log(`actual Old Row Count: ${actualOldRowCount}`);
      
      // oldEndRowを設定
      const oldEndRow = Math.min(desiredEndRow, actualOldRowCount - startRow + 1);
      console.log(`Old End Row: ${oldEndRow}`);
  
      const visibleOldRows: (string | number | boolean)[][] = [];
  
      // updateoldEndRow を数値で宣言
      const updateoldEndRow = startRow + oldEndRow - 1;
  
      // 指定された範囲の行を処理
      for (let i = startRow ; i < updateoldEndRow; i++) {
        const row = oldTableData[i];
        if (!tableOld.getRange().getRow(i).getHidden()) {
          visibleOldRows.push(row);
        }
      }
  
      console.log(`Processed ${visibleOldRows.length} visible rows from tableOld`);
      console.log(`Update Old End Row: ${updateoldEndRow}`);
  
      // データの内容を詳細表示
      console.log('Old Data:', JSON.stringify(visibleOldRows, null, 2));
  
      const result = {
        oldData: visibleOldRows,
        status: '処理終了',
        updateoldEndRow: updateoldEndRow,
        message: oldEndRow < desiredEndRow ? '最終回' : null
      };
  
      return JSON.stringify(result);
  
    } catch (error) {
      console.log(`Error processing old table: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw error; // エラーを再スローして呼び出し元で処理できるようにする
    }
}