/**
 * Dựng khung chạy thử cục bộ của module Report (không cần SharePoint):  node dev/build.js
 * Kết quả: temp/dev/index.html (mở bằng trình duyệt). Chi tiết: AGENTS.md mục "Chạy thử cục bộ".
 *  1. tsc biên dịch src/ + dev/ sang temp/dev/js
 *  2. nạp data_test/sharepoint_sample (template, ActivityP6.xlsx, logo, báo cáo mẫu; list CatalogItems, Report_Role từ các file *_list.xlsx ở docs/sharepoint_setup) thành seed.js
 *  3. webpack gộp, thay `core/services/SharePointClient` bằng bản giả lập trong bộ nhớ và `@microsoft/sp-http` bằng stub
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const webpack = require('webpack');

const root = path.resolve(__dirname, '..');
const out = path.join(root, 'temp', 'dev');
const js = path.join(out, 'js');
const sample = path.join(root, 'data_test', 'sharepoint_sample');

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

// 1. biên dịch
execFileSync(path.join(root, 'node_modules', '.bin', 'tsc'), ['-p', path.join(root, 'dev', 'tsconfig.json')], { stdio: 'inherit' });

// CSS không do tsc xử lý: chép sang cây biên dịch để import './x.css' tìm thấy
(function copyCss(dir) {
  fs.readdirSync(dir, { withFileTypes: true }).forEach((e) => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) copyCss(p);
    else if (e.name.endsWith('.css')) {
      const target = path.join(js, path.relative(root, p));
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.copyFileSync(p, target);
    }
  });
})(path.join(root, 'src'));

// 2. seed
const files = {};
(function walk(dir) {
  fs.readdirSync(dir, { withFileTypes: true }).forEach((e) => {
    const p = path.join(dir, e.name);
    const rel = path.relative(sample, p).split(path.sep).join('/');
    if (e.isDirectory()) walk(p);
    else if (!e.name.startsWith('.') && !e.name.startsWith('~$') && /^(_template\/.+\.json|ActivityP6\.xlsx|resource\/.+|data\/.+)$/.test(rel)) {
      files[rel] = fs.readFileSync(p).toString('base64');
    }
  });
})(sample);
// Nội dung list CatalogItems giả lập
const XLSX = require('xlsx');
const catalogRows = XLSX.utils.sheet_to_json(XLSX.readFile(path.join(root, 'docs', 'sharepoint_setup', 'CatalogItems_list.xlsx')).Sheets['CatalogItems'], { defval: '' });
// Phân quyền báo cáo: DisplayName nằm ở cột Title của list (như hướng dẫn tạo list)
const roleRows = XLSX.utils
  .sheet_to_json(XLSX.readFile(path.join(root, 'docs', 'sharepoint_setup', 'Report_Role_list.xlsx')).Sheets['Report_Role'], { defval: '' })
  .map((r) => ({ Title: r.DisplayName, Email: r.Email, Unit: r.Unit, DocID: r.DocID, DocTitle: r.DocTitle, Active: r.Active }));
fs.writeFileSync(path.join(out, 'seed.js'), `window.__SEED__=${JSON.stringify({ files, catalogRows, roleRows })};\n`);
fs.writeFileSync(
  path.join(out, 'index.html'),
  `<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Report – chạy thử cục bộ</title></head><body style="margin:0"><div id="root"></div><script src="seed.js"></script><script src="harness.js"></script></body></html>\n`
);

// 3. gộp
webpack(
  {
    mode: 'development',
    devtool: false,
    entry: path.join(js, 'dev', 'harness.js'),
    output: { path: out, filename: 'harness.js', chunkFilename: '[name].chunk.js', publicPath: '' },
    resolve: {
      extensions: ['.js', '.json'],
      modules: [path.join(root, 'node_modules')],
      alias: { '@microsoft/sp-http': path.join(js, 'dev', 'sp-http-stub.js') }
    },
    module: { rules: [{ test: /\.css$/, use: path.join(root, 'dev', 'style-loader.js') }] },
    plugins: [
      new webpack.NormalModuleReplacementPlugin(/(^|[\\/])SharePointClient$/, (r) => {
        r.request = path.join(js, 'dev', 'LocalSharePointClient.js');
      })
    ],
    performance: { hints: false }
  },
  (err, stats) => {
    if (err) throw err;
    console.log(stats.toString({ colors: false, modules: false, assets: true, chunks: false, children: false }));
    if (stats.hasErrors()) process.exit(1);
    console.log('\nXong. Mở:', path.join(out, 'index.html'));
  }
);
