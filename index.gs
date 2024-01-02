// シートに書き込み
function setInfo() {
    // シート取得
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = PropertiesService.getScriptProperties().getProperty("SHEET_NAME");
    var sheetData = book.getSheetByName(sheetName);

    // APIで取得
    var result = getSearchResult();

    if (result.length) {
        for (var i = 0; i < result.length; i++) {
            var title = '';
            var url = '';
            var snippet = '';

            if (isset(result[i]['title'])) {
                title = result[i]['title'];
            }
            if (isset(result[i]['link'])) {
                url = result[i]['link'];
            }
            if (isset(result[i]['snippet'])) {
                snippet = result[i]['snippet'];
            }

            // すでに登録されていないか判定（タイトルもしくはURL）
            if (!isExistsTitle(title, sheetData) && !isExistsUrl(url, sheetData)) {
                // 対象となるURL（市町村などの公式サイト）か判定
                // 対象となるタイトルか判定
                if (isTargetUrl(url) && isTargetTitle(title)) {
                    // 書き込み
                    var lastRow = sheetData.getLastRow();
                    sheetData.getRange(lastRow + 1, 1).setValue(title);
                    sheetData.getRange(lastRow + 1, 2).setValue(url);
                    sheetData.getRange(lastRow + 1, 3).setValue(snippet);
                }
            }
        }
    }
}
// 検索結果取得
function getSearchResult() {
    // API実行に必要な情報を取得

    // APIキー
    var apiKey = PropertiesService.getScriptProperties().getProperty("API_KEY");

    // 検索エンジンID
    var searchId = PropertiesService.getScriptProperties().getProperty("SEARCH_ENGINE_ID");

    // キーワード
    var query = PropertiesService.getScriptProperties().getProperty("QUERY");

    // 最大取得回数（API実行回数）
    var maxGetCount = 10;

    // 対象日付範囲（当日から2週間前まで）
    var today = new Date();
    var toYear = today.getFullYear();
    var toMonth = ('0' + (today.getMonth() + 1)).slice(-2);
    var toDay = ('0' + today.getDate()).slice(-2);

    today.setDate(today.getDate() - 14);
    var fromYear = today.getFullYear();
    var fromMonth = ('0' + (today.getMonth() + 1)).slice(-2);
    var fromDay = ('0' + today.getDate()).slice(-2);


    var fromDate = fromYear + fromMonth + fromDay;
    var toDate = toYear + toMonth + toDay;

    // API実行
    var items = [];
    var startNum = 1;
    for (var i1 = 1; i1 <= maxGetCount; i1++) {
        var apiUrl = "https://www.googleapis.com/customsearch/v1?";
        apiUrl += "key=";
        apiUrl += apiKey;
        apiUrl += "&cx=";
        apiUrl += searchId;
        apiUrl += "&sort=date:r:" + fromDate + ":" + toDate;
        apiUrl += "&start=";
        apiUrl += startNum;
        apiUrl += "&q=";
        apiUrl += query;

        var apiOptions = {
            method: 'get'
        };
        var result = UrlFetchApp.fetch(apiUrl, apiOptions);
        var resultJson = JSON.parse(result.getContentText());

        // 配列に格納
        if (isset(resultJson['items']) && resultJson['items'].length) {
            for (var i2 = 0; i2 < resultJson['items'].length; i2++) {
                items.push(resultJson["items"][i2]);
            }
        }

        // まだ取得できるか判定
        if (isset(resultJson['queries']['nextPage']) && resultJson['queries']['nextPage'].length) {
            if (isset(resultJson['queries']['nextPage'][0]['startIndex']) && resultJson['queries']['nextPage'][0]['startIndex']) {
                startNum = resultJson['queries']['nextPage'][0]['startIndex'];
            } else {
                break;
            }
        } else {
            break;
        }
    }
    return items;
}

// キー存在チェック
function isset(data) {
    if (data === "" || data === null || data === undefined) {
        return false;
    } else {
        return true;
    }
}

// すでに存在するか判定(URL)
function isExistsUrl(url, sheetData) {
    if (sheetData.getRange(3, 2, sheetData.getLastRow() - 2).getValues().flat().includes(url)) {
        return true;
    } else {
        return false;
    }
}

// すでに存在するか判定(タイトル)
function isExistsTitle(title, sheetData) {
    if (sheetData.getRange(3, 1, sheetData.getLastRow() - 2).getValues().flat().includes(title)) {
        return true;
    } else {
        return false;
    }
}

// 対象となるURLか判定（市町村などの公式サイト）
function isTargetUrl(url) {
    var regex = new RegExp('.*mlit.go.jp|.*pref.*jp|.*city.*jp|.*town.*jp|.*vill.*jp|.*lg.*jp|.*fire.*jp|.*119.*jp|.*shobo.*jp|.*shoubou.*jp', 'i');
    if (regex.test(url)) {
        return true;
    } else {
        return false;
    }
}

// 対象となるタイトルか判定（）
function isTargetTitle(title) {
    var eraName = PropertiesService.getScriptProperties().getProperty("ERA_NAME");
    var eraNameNum = PropertiesService.getScriptProperties().getProperty("ERA_NAME_NUM");

    var regex = new RegExp(eraName + '[2-' + (eraNameNum - 1) +']年', 'i');
    if (regex.test(title)) {
        return false;
    } else if (title.indexOf("平成") >= 0) {
        return false;
    } else {
        if (title.indexOf("しました") >= 0) {
            return false;
        } else {
            return true;
        }
    }
}