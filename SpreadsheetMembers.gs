/**
 * 管理スプレッドシート専用スクリプト
 * A列=区分、B列=表示名、C列=カレンダーID
 *
 * このファイルはWebアプリのCode.gsではなく、設定シート側の
 * 「拡張機能 > Apps Script」に配置してください。
 * メーリングリストの展開はメニュー実行時だけ行います。
 */

/**
 * シートを開いた際に専用メニューだけを追加（展開処理は実行しない）。
 * 既にonOpen()がある場合は、その関数の中にメニュー生成部分を移してください。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 メーリス管理')
    .addItem('メンバーを一括登録・更新', 'syncGroupsToMembers')
    .addToUi();
}

/**
 * 入力済みのメーリングリストを個人アドレスに展開する。
 * 表示名が空欄で、メールアドレスが入力されている行だけを対象とする。
 * グループ情報の取得にはGoogleグループ閲覧権限が必要。
 */
function syncGroupsToMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('メンバー');
  if (!sheet) {
    Browser.msgBox('「メンバー」シートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const newRows = [];
  const errorMessages = [];
  let expandedCount = 0;

  for (let row = 1; row < data.length; row++) {
    const groupVal = data[row][0] !== undefined ? String(data[row][0]).trim() : '';
    const nameVal = data[row][1] !== undefined ? String(data[row][1]).trim() : '';
    const rawEmail = data[row][2] !== undefined ? String(data[row][2]) : '';
    const emailVal = rawEmail.replace(/[\r\n\s ]+/g, '');

    // B列（表示名）が空欄の場合のみGoogleグループの取得を試みる
    if (emailVal && emailVal.includes('@') && nameVal === '') {
      let isGroupExpanded = false;

      try {
        const group = GroupsApp.getGroupByEmail(emailVal);
        if (group) {
          const users = group.getUsers();
          if (users && users.length > 0) {
            users.forEach(u => {
              const uEmail = u.getEmail();
              const uName = uEmail.split('@')[0];
              newRows.push([groupVal, uName, uEmail]);
            });
            isGroupExpanded = true;
            expandedCount++;
          } else {
            errorMessages.push('・' + emailVal + ' : グループ内にメンバーが存在しません');
          }
        }
      } catch (e) {
        const errorStr = e.toString();
        if (!errorStr.includes('Group not found') &&
            !errorStr.includes('グループが存在しません') &&
            !errorStr.includes('does not exist') &&
            !errorStr.includes('グループは見つかりません')) {
          errorMessages.push('・' + emailVal + ' : ' + errorStr);
        }
      }

      // 展開できないときは、元のアドレスを残す（既存動作を踏襲）
      // ※ B列に表示名を入れるため、次回の自動再試行はされない
      if (!isGroupExpanded) {
        newRows.push([groupVal, emailVal.split('@')[0], emailVal]);
      }
    } else {
      // 表示名のある行はそのまま保持
      newRows.push([data[row][0], data[row][1], data[row][2]]);
    }
  }

  if (newRows.length > 0) {
    // グループ展開で行数が増えても書き込み先が不足しないようにする
    const rowsNeeded = 1 + newRows.length;
    if (rowsNeeded > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), rowsNeeded - sheet.getMaxRows());
    }

    const existingRows = Math.max(sheet.getLastRow() - 1, 1);
    sheet.getRange(2, 1, existingRows, 3).clearContent();
    sheet.getRange(2, 1, newRows.length, 3).setValues(newRows);
  }

  if (errorMessages.length > 0) {
    Browser.msgBox('⚠️ 一部のアドレス展開でエラーが発生しました：\n\n' + errorMessages.join('\n'));
  } else if (expandedCount > 0) {
    Browser.msgBox('✅ 新規メーリス（' + expandedCount + '件）の展開が完了しました！');
  } else {
    Browser.msgBox('ℹ️ 新規に展開するメーリスはありませんでした。');
  }
}
