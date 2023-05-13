//
// 1列目(ISBN)を入力するとGoogle Books APIs を使用してタイトル等の情報をセルに出力する
//
// トリガー
// 実行する関数を選択: onEdit
// デプロイ時に実行: Head
// イベントのソースを選択: スプレッドシートから
// イベントの種類を選択: 編集時
//

function setBookInfo(){

  const sheet = SpreadsheetApp.getActiveSheet();            // 開いているシートのオブジェクトを取得
  const insertRow = sheet.getActiveCell().getRow();         // 選択している行のオブジェクトを取得
  const isbn = sheet.getActiveCell().getValue().toString(); // 選択しているセルの値を取得し、変数isbnに代入

  sheet.getRange(insertRow,2,1,8).setValue('');

  // ISBNの桁数をチェックしてISBN13ならISBN10へ変換
  let isbn10;
  if(10 == isbn.length) {
    isbn10 = isbn;
  } else if(13 == isbn.length) {
    isbn10 = convertToIsbn10(isbn); // ISBN13をISBN10へ
  } else {
    return;
  }

  // Google Books APIを叩いてレスポンスを取得  
  const response = UrlFetchApp.fetch('https://www.googleapis.com/books/v1/volumes?q=isbn:' + isbn + '&country=JP');
  const data = JSON.parse(response.getContentText());

  if(0 == data.totalItems) return;  // 検索結果が0なら処理終了

  const bookInfo = data.items[0].volumeInfo; // 本の情報を取得

  if('imageLinks' in bookInfo)
    if('thumbnail' in bookInfo.imageLinks)
      sheet.getRange(insertRow,2).setValue('=IMAGE("'+bookInfo.imageLinks.thumbnail+'")');            // サムネイル
  sheet.getRange(insertRow,3).setValue(bookInfo.title);                                               // タイトル
  sheet.getRange(insertRow,4).setValue('https://www.amazon.co.jp/dp/'+isbn10);                        // Amazonリンク
  if('authors'        in bookInfo) sheet.getRange(insertRow,5).setValue(bookInfo.authors.join());     // 著者
  if('publishedDate'  in bookInfo) sheet.getRange(insertRow,6).setValue(bookInfo.publishedDate);      // 発行日
  if('categories'     in bookInfo) sheet.getRange(insertRow,7).setValue(bookInfo.categories.join());  // カテゴリ

  const now = new Date();
  sheet.getRange(insertRow,8).setValue(now.toLocaleString('ja-JP'));                                  // 更新日時
  sheet.getRange(insertRow,9).setValue('=COUNTIF(A:A,"="&A'+insertRow+')');                           // 合計 ( 重複のチェック用 )
}

function convertToIsbn10(isbn13) {
    const sum = isbn13.split('').slice(3, 12).reduce((acc, c, i) => {
        return acc + (c[0] - '0') * (10 - i);
    }, 0);

    let isbn10 = isbn13.substring(3, 12);
    const checkDigit = 11 - sum % 11;
    if(10 == checkDigit) {
      return isbn10 + 'X';
    } else if(11 == checkDigit) {
      return isbn10 + '0';
    } else {
      return isbn10 + checkDigit.toString();
    }
}

function onEdit(e){
  if(e.range.getColumn() == 1){
    setBookInfo();
  }
}