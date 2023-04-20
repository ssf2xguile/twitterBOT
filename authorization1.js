// OAuth1.0の認証で利用する値を変数にセット
var accessTokenUrl = "https://api.twitter.com/oauth/access_token";
var requestTokenUrl = "https://api.twitter.com/oauth/request_token";
var authorizationUrl = "https://api.twitter.com/oauth/authorize";
var consumerKey = "your consumerKey"; // Consumer Key をセット
var consumerSecret = "your consumerSecret"; // Consumer Secret をセット
var serviceName = 'twitter test';

// OAuth1.0の認証で、Twitterにアクセスする関数
function getService() {
  return OAuth1.createService(serviceName)
    // OAuth1.0の認証で利用する値をセット
    .setAccessTokenUrl(accessTokenUrl)
    .setRequestTokenUrl(requestTokenUrl)
    .setAuthorizationUrl(authorizationUrl)
    .setConsumerKey(consumerKey)
    .setConsumerSecret(consumerSecret)
    // 認証の確認後に実行するコールバック関数を指定
    .setCallbackFunction('authCallback')
    // 生成したトークンを、GASのプロパティストアに保存（永続化）
    .setPropertyStore(PropertiesService.getUserProperties());
}


// 認証の確認後に表示する可否メッセージを指定する関数
function authCallback(request) {
  const service = getService();
  const isAuthorized = service.handleCallback(request);

  if (isAuthorized) {
    return HtmlService.createHtmlOutput('認証が許可されました。');
  } else {
    return HtmlService.createHtmlOutput('認証が拒否されました。');
  }
}


// OAuth1.0のトークンを生成し、認証画面のURLを表示する関数（関数はスプレッドシートの画面で利用）
function authorizeLink() {
  const ui = SpreadsheetApp.getUi();
  const service = getService();

  if (!service.hasAccess()) {
    const authorizationUrl = service.authorize(); // トークンを生成し、認証ページのURLを返します
    const template = HtmlService.createTemplate('<a href="<?= authorizationUrl ?>" target="_blank">Authorize Link</a>');
    template.authorizationUrl = authorizationUrl;
    const output = template.evaluate();
    ui.showModalDialog(output, 'OAuth1.0認証');
  } else {
     ui.alert("OAuth1.0認証はすでに許可されています。");
  }
}


// プロパティストアに保存したトークンをリセットする関数
function clearService(){
  OAuth1.createService(serviceName)
    .setPropertyStore(PropertiesService.getUserProperties())
    .reset();
}


// API リクエストを送信するための関数
function makeRequest(url, options) {
  const service = getService();
  const res = service.fetch(url, options);
  const result = JSON.parse(res.getContentText());

  // console.log(JSON.stringify(result));
  return result
}
