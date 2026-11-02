// V9：起動直後から CSP 違反を記録する（office.js の読み込みより前に実行される）
window.__cspViolations = [];
document.addEventListener('securitypolicyviolation', function (e) {
  window.__cspViolations.push({
    directive: e.effectiveDirective || e.violatedDirective,
    blockedURI: e.blockedURI,
    source: e.sourceFile ? e.sourceFile + ':' + e.lineNumber : '',
    at: new Date().toISOString(),
  });
  window.dispatchEvent(new Event('csp-violation'));
});
