/**
 * 項目名の名寄せを行う。キーワードマッチングまたは不要な英数字の除去。
 * @param {string} name - 元の項目名
 * @returns {string} 名寄せ後の項目名
 * @private
 */
function normalizeItemName_(name) {
  // 1. 全角英数字を半角化し、全角空白を半角空白に置換、さらに大文字へ統一
  const n = name
    .replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xfee0)) // 英数字半角化
    .replace(/\u3000/g, ' ') // 全角空白を半角へ【追加】
    .toUpperCase();

  if (n.includes('1PASSWORD')) return '1Password';
  if (n.includes('DOCKER, INC.')) return 'Docker';
  if (n.includes('DRI*PVTLTRACKER')) return 'Pivotal Tracker';
  if (n.includes('MAILTRAP')) return 'Mailtrap';
  if (n.includes('AMAZON WEB SERVICES')) return 'Amazon Web Services';
  if (n.includes('PAPERTRAIL-SOLARWINDS') || n.includes('SOLARWINDS'))
    return 'Papertrail';
  if (n.includes('PULUMI CORPORATION')) return 'Pulumi';
  if (n.includes('ROLLBAR')) return 'Rollbar';
  if (n.includes('WWW.DEEPL.COM') || n.includes('DEEPL')) return 'DeepL';
  if (n.includes('ZOOM')) return 'Zoom';
  if (n.includes('DROPBOX')) return 'Dropbox';
  if (n.includes('GOOGLE*WORKSPACE') || n.includes('GOOGLE*GSUITE'))
    return 'Google Workspace';
  if (n.includes('AMAZON')) return 'Amazon';
  if (n.includes('GITHUB')) return 'GitHub';
  if (n.includes('さくらインターネット')) return 'さくらインターネット';
  if (n.includes('OPENAI')) return 'OpenAI';
  if (n.includes('HEROKU')) return 'Heroku';
  if (n.includes('CODE CLIMATE')) return 'Code Climate';
  if (
    n.includes('カブシキガイシャボックス') ||
    n.includes('カブシキガイシヤボツクス')
  )
    return 'Box';
  if (n.includes('SKYPE')) return 'Skype';
  if (n.includes('オナマエドツトコムドメイン')) return 'お名前.COMドメイン';
  if (n.includes('LINEAR.APP')) return 'LINEAR.APP';
  if (n.includes('CLAUDE.AI')) return 'Claude.AI';
  if (n.includes('ビジ得チャンス')) return 'ビジ得チャンス';

  return n;
}
