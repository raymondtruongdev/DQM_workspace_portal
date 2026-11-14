// Loader tối giản cho khung chạy thử: nhúng nội dung file .css vào <style> (CSS của dự án không dùng @import/url()).
module.exports = function styleInject(source) {
  return `var s=document.createElement('style');s.setAttribute('data-src',${JSON.stringify(this.resourcePath.split('/').slice(-2).join('/'))});s.textContent=${JSON.stringify(source)};document.head.appendChild(s);`;
};
