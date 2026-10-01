(function () {
  var page = document.querySelector('.page');
  var NOME = page.getAttribute('data-estudo') || 'estudo';
  function read(key) { try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch (e) { return null; } }
  function write(key, val) { try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) {} }

  /* ---------- Checklist ---------- */
  var boxes = Array.prototype.slice.call(document.querySelectorAll('#checklist input'));
  var count = document.getElementById('ck-count');
  var total = document.getElementById('ck-total');
  if (boxes.length && count) {
    var CK = 'av-' + NOME + '-checklist';
    var saved = read(CK) || {};
    if (total) { total.textContent = boxes.length; }
    var refreshChecks = function () {
      var n = 0, state = {};
      boxes.forEach(function (b, i) { if (b.checked) { n++; } state[b.id || i] = b.checked; });
      count.textContent = n;
      return state;
    };
    boxes.forEach(function (b, i) {
      if (saved[b.id || i]) { b.checked = true; }
      b.addEventListener('change', function () { write(CK, refreshChecks()); });
    });
    refreshChecks();
    var ckReset = document.getElementById('ck-reset');
    if (ckReset) {
      ckReset.addEventListener('click', function () {
        boxes.forEach(function (b) { b.checked = false; });
        write(CK, refreshChecks());
      });
    }
  }

  /* ---------- Exercício ---------- */
  /* Os dados vêm do bloco JSON de id quiz-data:
     { "opcoes": {"A": "Nome da opção A", ...}, "itens": [ {"t": "frase", "a": "A", "w": "por quê"} ] } */
  var quiz = document.getElementById('quiz');
  var dataEl = document.getElementById('quiz-data');
  if (quiz && dataEl) {
    var data = JSON.parse(dataEl.textContent);
    var NAMES = data.opcoes, ITEMS = data.itens, KEYS = Object.keys(NAMES);
    var score = document.getElementById('qz-score');
    var done = document.getElementById('qz-done');
    document.getElementById('qz-total').textContent = ITEMS.length;
    var hits = 0, answered = 0;
    var render = function () {
      hits = 0; answered = 0; score.textContent = 0; done.textContent = 0;
      quiz.textContent = '';
      ITEMS.forEach(function (item, idx) {
        var row = document.createElement('div'); row.className = 'q';
        var p = document.createElement('p'); p.textContent = (idx + 1) + '. ' + item.t;
        var opts = document.createElement('div'); opts.className = 'opts';
        opts.setAttribute('role', 'group'); opts.setAttribute('aria-label', 'Resposta do item ' + (idx + 1));
        var why = document.createElement('p'); why.className = 'why'; why.hidden = true;
        KEYS.forEach(function (k) {
          var b = document.createElement('button');
          b.type = 'button'; b.textContent = k;
          b.setAttribute('aria-label', NAMES[k]); b.title = NAMES[k];
          b.addEventListener('click', function () {
            var ok = k === item.a;
            Array.prototype.forEach.call(opts.children, function (o) {
              o.disabled = true;
              if (o.textContent === item.a) { o.className = 'right'; }
            });
            if (!ok) { b.className = 'wrong'; }
            var lead = document.createElement('b');
            lead.textContent = (ok ? 'Certo. ' : 'A resposta é ' + item.a + (NAMES[item.a] !== item.a ? ', ' + NAMES[item.a] : '') + '. ');
            why.textContent = ''; why.appendChild(lead); why.appendChild(document.createTextNode(item.w));
            why.className = 'why ' + (ok ? 'ok' : 'no'); why.hidden = false;
            answered++; if (ok) { hits++; }
            score.textContent = hits; done.textContent = answered;
          });
          opts.appendChild(b);
        });
        row.appendChild(p); row.appendChild(opts); row.appendChild(why);
        quiz.appendChild(row);
      });
    };
    render();
    document.getElementById('qz-reset').addEventListener('click', render);
  }
})();
