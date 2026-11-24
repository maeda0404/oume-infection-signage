const DATA_URL = 'data/status.json';

const LEVELS = [
  'level-normal',
  'level-epidemic',
  'level-advisory',
  'level-warning'
];


/* =========================================================
   数値変換
========================================================= */

function n(value) {
  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return null;
  }

  const x = Number(value);

  return Number.isFinite(x) ? x : null;
}


/* =========================================================
   流行レベル判定

   1未満        非流行
   1以上10未満  流行中
   10以上30未満 注意報レベル
   30以上       警報レベル
========================================================= */

function classify(value) {
  if (value >= 30) {
    return [
      '警報レベル',
      'level-warning'
    ];
  }

  if (value >= 10) {
    return [
      '注意報レベル',
      'level-advisory'
    ];
  }

  if (value >= 1) {
    return [
      '流行中',
      'level-epidemic'
    ];
  }

  return [
    '非流行',
    'level-normal'
  ];
}


/* =========================================================
   週番号表示

   JSONが
   "week": "第46週"
   でも
   "weekNumber": 46
   でも対応
========================================================= */

function formatWeek(item) {
  if (!item) {
    return '--';
  }

  if (item.week) {
    const text = String(item.week);

    if (text.includes('週')) {
      return text;
    }

    return `第${text}週`;
  }

  const weekNumber = n(item.weekNumber);

  if (weekNumber !== null) {
    return `第${weekNumber}週`;
  }

  return '--';
}


/* =========================================================
   増減表示
========================================================= */

function renderTrend(element, current, previous, difference) {
  if (!element) {
    return;
  }

  if (
    current === null ||
    previous === null
  ) {
    element.textContent = '前週 --';
    return;
  }

  let diff = n(difference);

  /*
    differenceがJSONにない場合は
    最新値 - 前週値で算出
  */
  if (diff === null) {
    diff = current - previous;
  }

  const arrow =
    diff > 0
      ? '↑'
      : diff < 0
        ? '↓'
        : '→';

  const cssClass =
    diff > 0
      ? 'up'
      : diff < 0
        ? 'down'
        : 'same';

  const sign =
    diff > 0
      ? '+'
      : '';

  element.innerHTML =
    `前週 ${previous.toFixed(2)} ` +
    `<span class="${cssClass}">` +
    `${arrow} ${sign}${diff.toFixed(2)}` +
    `</span>`;
}


/* =========================================================
   1枚分の感染症カードを描画

   item      = 西多摩データ
   tokyoItem = 東京都データ
========================================================= */

function render(prefix, item, tokyoItem) {
  const card =
    document.getElementById(`${prefix}Card`);

  const badge =
    document.getElementById(`${prefix}Level`);

  const value =
    document.getElementById(`${prefix}Value`);

  const trend =
    document.getElementById(`${prefix}Trend`);

  const week =
    document.getElementById(`${prefix}Week`);

  const tokyo =
    document.getElementById(`${prefix}Tokyo`);


  if (
    !card ||
    !badge ||
    !value ||
    !trend ||
    !week ||
    !tokyo
  ) {
    console.error(
      `表示要素が見つかりません: ${prefix}`
    );

    return;
  }


  /*
    新しいstatus.jsonでは

    perSentinel
    previousPerSentinel
    difference

    を使用
  */

  const current =
    n(item?.perSentinel);

  const previous =
    n(item?.previousPerSentinel);

  const difference =
    n(item?.difference);

  const tokyoValue =
    n(tokyoItem?.perSentinel);


  /* -------------------------------------------------------
     西多摩データが取得できない場合
  ------------------------------------------------------- */

  if (current === null) {
    card.classList.remove(...LEVELS);

    badge.textContent = '確認中';
    value.textContent = '--';
    trend.textContent = '前週 --';
    week.textContent = '--';

    tokyo.textContent =
      tokyoValue === null
        ? '--'
        : tokyoValue.toFixed(2);

    return;
  }


  /* -------------------------------------------------------
     レベル

     Python側のlevelを優先。
     存在しない場合のみJS側で判定。
  ------------------------------------------------------- */

  const calculated =
    classify(current);

  const levelLabel =
    item?.level || calculated[0];

  const levelClass =
    calculated[1];


  card.classList.remove(...LEVELS);
  card.classList.add(levelClass);

  badge.textContent =
    levelLabel;


  /* -------------------------------------------------------
     最新値
  ------------------------------------------------------- */

  value.textContent =
    current.toFixed(2);


  /* -------------------------------------------------------
     前週・増減
  ------------------------------------------------------- */

  renderTrend(
    trend,
    current,
    previous,
    difference
  );


  /* -------------------------------------------------------
     週番号
  ------------------------------------------------------- */

  week.textContent =
    formatWeek(item);


  /* -------------------------------------------------------
     東京都
  ------------------------------------------------------- */

  tokyo.textContent =
    tokyoValue === null
      ? '--'
      : tokyoValue.toFixed(2);
}


/* =========================================================
   status.json取得
========================================================= */

async function load() {
  try {
    /*
      GitHub Pagesのキャッシュ対策として
      時刻をクエリパラメーターに付加
    */

    const response = await fetch(
      `${DATA_URL}?t=${Date.now()}`,
      {
        cache: 'no-store'
      }
    );


    if (!response.ok) {
      throw new Error(
        `status.json HTTP ${response.status}`
      );
    }


    const data =
      await response.json();


    console.log(
      'status.json:',
      data
    );


    /* =====================================================
       西多摩 インフルエンザ
    ===================================================== */

    render(
      'influenza',
      data?.nishitama?.influenza,
      data?.tokyo?.influenza
    );


    /* =====================================================
       西多摩 新型コロナ
    ===================================================== */

    render(
      'covid19',
      data?.nishitama?.covid19,
      data?.tokyo?.covid19
    );


    /* =====================================================
       データ状態表示
    ===================================================== */

    const statusElement =
      document.getElementById(
        'dataStatus'
      );


    if (statusElement) {
      if (
        data?.sourceStatus === '正常'
      ) {
        statusElement.textContent =
          'データ更新済み';
      } else if (
        data?.sourceStatus
      ) {
        statusElement.textContent =
          data.sourceStatus;
      } else {
        statusElement.textContent =
          'データ更新済み';
      }
    }


  } catch (error) {
    console.error(
      'status.json取得エラー:',
      error
    );


    const statusElement =
      document.getElementById(
        'dataStatus'
      );


    if (statusElement) {
      statusElement.textContent =
        '更新確認中';
    }
  }
}


/* =========================================================
   初回読み込み
========================================================= */

load();


/* =========================================================
   10分ごとに再取得
========================================================= */

setInterval(
  load,
  10 * 60 * 1000
);
