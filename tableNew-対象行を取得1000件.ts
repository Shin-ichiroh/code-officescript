async function main(workbook: ExcelScript.Workbook, startRow: number) {
    try {
      const tableNew = workbook.getTable("tableNew");
      const desiredEndRow = 1000;
      if (!tableNew) {
        throw new Error("Table 'tableNew' not found");
      }
  
      // テーブルのデータを一度に取得
    //  const newTableData = tableNew.getRangeBetweenHeaderAndTotal().getValues();
      const newTableData = tableNew.getRange().getValues();
      const actualNewRowCount = newTableData.length;
      console.log(`actual New Row Count: ${actualNewRowCount}`);
      
      // newEndRowを設定
      const newEndRow = Math.min(desiredEndRow, actualNewRowCount - startRow + 1);    
      console.log(`New End Row: ${newEndRow}`);
      
      const visibleNewRows: (string | number | boolean)[][] = [];
  
      // updatenewEndRow を数値で宣言
      const updatenewEndRow = startRow + newEndRow - 1;
  
      // 指定された範囲の行を処理
      for (let i = startRow; i < updatenewEndRow; i++) {
        const row = newTableData[i];
        if (!tableNew.getRange().getRow(i).getHidden()) {
          visibleNewRows.push(row);
        }
      }
  
      console.log(`Processed ${visibleNewRows.length} visible rows from tableNew`);
      console.log(`Update New End Row: ${updatenewEndRow}`);
  
      // データの内容を詳細表示
      console.log('New Data:', JSON.stringify(visibleNewRows, null, 2));
  
      const result = {
        newData: visibleNewRows,
        status: '処理終了',
        updatenewEndRow: updatenewEndRow,
        message: newEndRow < desiredEndRow ? '最終回' : null
      };
  
      return JSON.stringify(result);
  
    } catch (error) {
      console.log(`Error processing new table: ${error instanceof Error ? error.message : 'Unknown error'}`);
      throw error; // エラーを再スローして呼び出し元で処理できるようにする
    }
}