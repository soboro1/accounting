(function () {
  // ===== 日程データ(更新するときはここを書き換える) =====
  var EXAMS = [
    { no: 174, date: '2026-11-15', grades: [1, 2, 3], apply: { start: '2026-09-28', end: '2026-10-19' } },
    { no: 175, date: '2027-02-28', grades: [2, 3] },
    { no: 176, date: null, approx: '2027年6月(予定)', grades: [1, 2, 3] }
  ];

  var $ = function (id) { return document.getElementById(id); };
  var DOW = '日月火水木金土';

  function parse(s) {
    var p = s.split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  function today() {
    var n = new Date();
    return new Date(n.getFullYear(), n.getMonth(), n.getDate());
  }
  function diff(a, b) { return Math.round((b - a) / 86400000); }
  function fmt(d) { return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日(' + DOW[d.getDay()] + ')'; }

  var grade = 3;
  try { grade = +localStorage.getItem('bokiCountdownMini') || 3; } catch (e) {}

  function nextExam(t) {
    for (var i = 0; i < EXAMS.length; i++) {
      var e = EXAMS[i];
      if (e.grades.indexOf(grade) >= 0 && (!e.date || diff(t, parse(e.date)) >= 0)) return e;
    }
    return null;
  }

  function setBadge(cls, text) {
    $('badge').className = 'badge ' + cls;
    $('badge').innerHTML = text;
  }

  function render() {
    var t = today();
    Array.prototype.forEach.call(document.querySelectorAll('.chip'), function (b) {
      b.classList.toggle('on', +b.getAttribute('data-grade') === grade);
    });

    var e = nextExam(t);
    $('count').style.display = 'none';
    $('msg').style.display = 'none';
    if (!e) {
      $('exam').textContent = grade + '級';
      $('msg').textContent = '日程未発表';
      $('msg').style.display = 'block';
      setBadge('wait', '申込:未発表');
      return;
    }

    $('exam').innerHTML = '<span>第' + e.no + '回 統一試験・' + grade + '級</span> <span>' + (e.date ? fmt(parse(e.date)) : e.approx) + '</span>';
    if (!e.date) {
      $('msg').textContent = '日程未発表';
      $('msg').style.display = 'block';
    } else {
      var n = diff(t, parse(e.date));
      if (n === 0) {
        $('msg').textContent = '今日が試験日!';
        $('msg').style.display = 'block';
      } else {
        $('num').textContent = n;
        $('count').style.display = 'flex';
      }
    }

    if (e.apply) {
      var st = parse(e.apply.start), en = parse(e.apply.end);
      if (t < st) setBadge('wait', '申込:受付前');
      else if (t <= en) setBadge('open', '申込受付中<span class="long">(締切は地域ごと)</span>');
      else setBadge('closed', '申込:多くの地域で終了');
    } else {
      setBadge('wait', '申込:未発表');
    }
  }

  Array.prototype.forEach.call(document.querySelectorAll('.chip'), function (b) {
    b.addEventListener('click', function () {
      grade = +b.getAttribute('data-grade');
      try { localStorage.setItem('bokiCountdownMini', grade); } catch (e) {}
      render();
    });
  });

  render();
})();
