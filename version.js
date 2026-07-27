/* GESMS 단일 버전 관리 파일: 새 버전 배포 시 아래 한 줄만 변경 */
window.GESMS_APP_VERSION = 'V10.9.2 Stable';

(function applyGesmsVersion(){
  const version = window.GESMS_APP_VERSION;
  const replaceVersion = (text) => typeof text === 'string'
    ? text.replace(/V\d+(?:\.\d+){1,3}/gi, version)
    : text;

  const run = () => {
    document.title = replaceVersion(document.title);
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach(node => {
      const next = replaceVersion(node.nodeValue);
      if (next !== node.nodeValue) node.nodeValue = next;
    });
    document.querySelectorAll('[data-gesms-version]').forEach(el => {
      el.textContent = version;
    });
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', run, { once: true });
  } else {
    run();
  }
})();
