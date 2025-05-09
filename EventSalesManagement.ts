function main(workbook: ExcelScript.Workbook) {
    // 日付の取得
    let now = new Date();
    // let formattedDate = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}_${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}`;

     let formattedDate = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;

    // シートの設定
    let motoSheet = workbook.getWorksheet("Export");
    let nayoseSheet = workbook.getWorksheet(formattedDate);
    
    if (nayoseSheet) {
        console.log(`シート "${formattedDate}" は既に存在します。処理をスキップします。`);
        return;
    }
    
    nayoseSheet = workbook.addWorksheet(formattedDate);
    
    // 実際の最終行を取得
    let motoLastRow = motoSheet.getUsedRange().getRowCount();
    console.log("データの最終行:", motoLastRow);
    let motoRowValues = motoSheet.getRange(`A${motoLastRow}:D${motoLastRow}`).getValues()[0];
    let conditionA = motoRowValues[0] != "" && motoRowValues[0] != null;
    let conditionD = motoRowValues[3] == "" || motoRowValues[3] == null;
    console.log("条件A:", conditionA);      // 集計行の場合でもTrue、何もなければFalse
    console.log("条件D:", conditionD);      // 正しい行であればFalse、消してよい行であればTrue

    // 不要な行の場合は行削除
    if (conditionA && conditionD) {
        console.log("条件を満たしています。この行を削除します。");
        motoSheet.getRange(`A${motoLastRow}:R${motoLastRow}`).delete(ExcelScript.DeleteShiftDirection.up);
        console.log("行を削除しました。");
    } else {
        console.log("条件を満たしていません。行は削除されません。");
    }

    // 実際の最終行を取得２（不要な行が２つあるのでもう１回実行する）
    motoLastRow = motoSheet.getUsedRange().getRowCount();
    console.log("データの最終行:", motoLastRow);
    motoRowValues = motoSheet.getRange(`A${motoLastRow}:D${motoLastRow}`).getValues()[0];
    conditionA = motoRowValues[0] != "" && motoRowValues[0] != null;
    conditionD = motoRowValues[3] == "" || motoRowValues[3] == null;
    console.log("条件A:", conditionA);      // 集計行の場合でもTrue、何もなければFalse
    console.log("条件D:", conditionD);      // 正しい行であればFalse、消してよい行であればTrue

    // 不要な行の場合は行削除２
    if (conditionA && conditionD) {
        console.log("条件を満たしています。この行を削除します。");
        motoSheet.getRange(`A${motoLastRow}:R${motoLastRow}`).delete(ExcelScript.DeleteShiftDirection.up);
        console.log("行を削除しました。");
    } else {
        console.log("条件を満たしていません。行は削除されません。");
    }

    // 最終的なデータの最終行を取得
    motoLastRow = motoSheet.getUsedRange().getRowCount();
    console.log("最終的なデータの最終行:", motoLastRow);

    // データのコピー
    let headerValues = motoSheet.getRange("A1:R1").getValues();
    nayoseSheet.getRange("A1:R1").setValues(headerValues);

    let dataValues = motoSheet.getRange(`A2:E${motoLastRow}`).getValues();
    nayoseSheet.getRange(`A2:E${motoLastRow}`).setValues(dataValues);

    // 重複の削除
    nayoseSheet.getRange("A:E").removeDuplicates([0, 1, 2, 3, 4], true);
    let nayoseLastRow = nayoseSheet.getUsedRange().getUsedRange().getRowCount();

    // データの取得
    let rngData1 = nayoseSheet.getRange(`A2:R${nayoseLastRow}`).getValues();
    let rngData2 = motoSheet.getRange(`A2:R${motoLastRow}`).getValues();

    console.log("rngData1の行数:", rngData1.length);
    console.log("rngData2の行数:", rngData2.length);

    // データの処理
    let amount: number = 0;
    let quantity: number = 0;

    for (let i = 0; i < rngData1.length; i++) {
        for (let j = 0; j < rngData2.length; j++) {
            if (rngData1[i][0] == rngData2[j][0] && rngData1[i][1] == rngData2[j][1] &&
                rngData1[i][2] == rngData2[j][2] && rngData1[i][3] == rngData2[j][3] &&
                rngData1[i][4] == rngData2[j][4]) {
                for (let k = 5; k < 16; k++) {
                    if (rngData2[j][k] !== "") {
                        rngData1[i][k] = rngData2[j][k];
                    }
                }
                amount += Number(rngData2[j][16]);
                quantity += Number(rngData2[j][17]);
            }
        }
        rngData1[i][16] = amount;
        amount = 0;
        rngData1[i][17] = quantity;
        quantity = 0;
    }

    // 処理したデータを書き戻す
    nayoseSheet.getRange(`A2:R${nayoseLastRow}`).setValues(rngData1);

    // 列幅の自動調整
    nayoseSheet.getRange("A:R").getFormat().autofitColumns();

    // テーブルの作成
    // let selectedSheet = workbook.getWorksheets()[0];
    let range = nayoseSheet.getUsedRange();
    console.log(range.getRowCount())
    let table = nayoseSheet.addTable(range.getAddress(), true);
    table.setName("table1");
    table.setPredefinedTableStyle("TableStyleMedium2");


    // ファイルの保存（注：この部分はExcel on the webでは動作しない可能性があります）
    console.log(`ファイル名: イベント参加企業${formattedDate}`);
}
