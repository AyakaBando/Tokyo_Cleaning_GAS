function getSettings() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Settings');
    const values = sheet.getDataRange().getValues();

    const settings = {};
    for (let i = 0; i < values.length; i++) {
        const key = values[i][0];
        const value = values[i][1];
        if (key) settings[key] = value;
    }

    return settings;
}

function getMembers() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Members');
    const values = sheet.getDataRange().getValues();
    const headers = values.shift();

    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const members = values.filter(r => String(r[idx.active]).toUpperCase() === 'TRUE').map(r => ({
        no: Number(r[idx.no]),
        name: r[idx.name],
        slack_id: r[idx.slack_id],
        active: r[idx.active]
    })).sort((a, b) => a.no - b.no);

    return members;
}

function normalizeDate(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays(date, days) {
    const d = new Date(date);
    d.setDate(d.getDate() + days);
    return d;
}

function diffWeeks(baseDate, targetDate) {
    const ms = normalizeDate(targetDate) - normalizeDate(baseDate);
    return Math.floor(ms / (1000 * 60 * 60 * 24 * 7));
}

function getMondaysOfMonth(year, month) {
    const dates = [];
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);

    let d = new Date(firstDay);
    while (d.getDay() !== 1) {
        d.setDate(d.getDate() + 1);
    }

    while (d <= lastDay) {
        dates.push(new Date(d));
        d.setDate(d.getDate() + 7);
    }

    return dates;
}

function getSkipWeeks() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('SkipWeeks');
    if (!sheet) return new Set();

    const values = sheet.getDataRange().getValues();
    const headers = values.shift();

    const idx = {};
    headers.forEach((h, i) => idx[h] = i);

    const skipWeeks = new Set();
    const tz = Session.getScriptTimeZone();

    values.forEach(r => {
        const active = String(r[idx.active]).toUpperCase() === 'TRUE';
        if (!active) return;

        let weekStart = r[idx.week_start];
        if (!weekStart) return;

        if (!(weekStart instanceof Date)) {
            weekStart = new Date(weekStart);
        }

        if (isNaN(weekStart.getTime())) return;

        const key = Utilities.formatDate(
            new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate()),
            tz,
            'yyyy/MM/dd'
        );
        skipWeeks.add(key);
    });

    Logger.log([...skipWeeks]);
    return skipWeeks;
}

function countWorkedWeeks(baseMonday, targetWeekStart, skipWeeks) {
    let count = 0;
    let d = new Date(baseMonday);

    while (d < targetWeekStart) {
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        if (!skipWeeks.has(key)) {
            count++;
        }
        d.setDate(d.getDate() + 7);
    }

    return count;
}

function getDutyMembersForWeek(weekStart, members, dutyCount, baseMonday, baseStartNo, skipWeeks) {
    const memberCount = members.length;
    if (memberCount === 0) return [];

    const weekKey = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    if (skipWeeks.has(weekKey)) {
        return [];
    }

    const actualDutyCount = Math.min(Number(dutyCount), memberCount);
    const baseIndex = members.findIndex(m => m.no === Number(baseStartNo));
    if (baseIndex === -1) {
        throw new Error(`base_start_no does not exist in members sheet. base_start_no: ${baseStartNo}`);
    }

    const workedWeeks = countWorkedWeeks(baseMonday, weekStart, skipWeeks);

    const startIndex =
        ((baseIndex + (workedWeeks * actualDutyCount)) % memberCount + memberCount) % memberCount;

    const result = [];
    for (let i = 0; i < actualDutyCount; i++) {
        result.push(members[(startIndex + i) % memberCount]);
    }

    return result;
}

function buildNextMonthSchedule() {
    const settings = getSettings();
    const members = getMembers();
    const skipWeeks = getSkipWeeks();

    const dutyCount = Number(settings.duty_count || 0);
    const baseMonday = normalizeDate(new Date(settings.base_monday));
    const baseStartNo = Number(settings.base_start_no);

    const today = new Date();
    const nextMondayDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const year = nextMondayDate.getFullYear();
    const month = nextMondayDate.getMonth() + 1;

    const mondays = getMondaysOfMonth(year, month);

    const rows = [];
    mondays.forEach(weekStart => {
        const dutyMembers = getDutyMembersForWeek(weekStart, members, dutyCount, baseMonday, baseStartNo, skipWeeks);

        dutyMembers.forEach(member => {
            rows.push([
                weekStart,
                member.no,
                member.name,
                member.slack_id
            ]);
        });
    });

    return {
        year,
        month,
        rows,
    };
}

function outputNextMonthSchedule() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Schedule');
    if (!sheet) {
        sheet = ss.insertSheet('Schedule');
    }

    sheet.clearContents();

    const result = buildNextMonthSchedule();
    const header = [['week_start', 'no', 'name', 'slack_id']];
    sheet.getRange(1, 1, 1, header[0].length).setValues(header);

    if (result.rows.length > 0) {
        sheet.getRange(2, 1, result.rows.length, result.rows[0].length).setValues(result.rows);
        sheet.getRange(2, 1, result.rows.length, 1).setNumberFormat('yyyy/MM/dd');
    }
}

function sendNextMonthCleaningSummaryToSlack() {
    const settings = getSettings();
    const webhookUrl = settings.webhook_url;
    if (!webhookUrl) {
        throw new Error('Settings シートに webhook_url がありません。');
    }

    const members = getMembers();
    const skipWeeks = getSkipWeeks();

    const dutyCount = Number(settings.duty_count || 8);
    const baseMonday = normalizeDate(new Date(settings.base_monday));
    const baseStartNo = Number(settings.base_start_no);

    const today = new Date();
    const nextMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const year = nextMonthDate.getFullYear();
    const month = nextMonthDate.getMonth() + 1;

    const mondays = getMondaysOfMonth(year, month);
    const tz = Session.getScriptTimeZone();

    const lines = [];
    lines.push(`【${year}年${month}月 掃除当番一覧】`);
    lines.push('');

    mondays.forEach(weekStart => {
        const weekKey = Utilities.formatDate(weekStart, tz, 'yyyy/MM/dd');
        const label = `*【${weekStart.getMonth() + 1}/${weekStart.getDate()}週】*`;

        if (skipWeeks.has(weekKey)) {
            lines.push(`${label}\n 長期休暇のため当番なし\n`);
            return;
        }

        const dutyMembers = getDutyMembersForWeek(
            weekStart,
            members,
            dutyCount,
            baseMonday,
            baseStartNo,
            skipWeeks
        );

        const names = dutyMembers.map(m => m.slack_id ? m.slack_id : m.name).join('・');
        lines.push(`${label}\n ${names}\n`);
    });

    const payload = {
        text: lines.join('\n')
    };

    UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
    });
}

function sendThisWeekCleaningToSlack() {
    const settings = getSettings();
    const webhookUrl = settings.webhook_url;
    if (!webhookUrl) {
        throw new Error('There is no webhook URL on Settings Sheet')
    }

    const members = getMembers();
    const skipWeeks = getSkipWeeks();
    const holidaySet = getCabinetOfficeHolidaySet();

    const dutyCount = Number(settings.duty_count || 8);
    const baseMonday = normalizeDate(new Date(settings.base_monday));
    const baseStartNo = Number(settings.base_start_no);

    const today = new Date();
    const tz = Session.getScriptTimeZone();

    const todayDay = today.getDay();
    // if (todayDay !== 1 && todayDay !== 2) {
    //     Logger.log('月曜・火曜以外のため送信しません。');
    //     return;
    // }

    // // 月曜が祝日なら月曜は送らない
    // if (todayDay === 1 && isJapaneseHoliday(today, holidaySet)) {
    //     Logger.log('月曜日が祝日のため、本日は送信しません。');
    //     return;
    // }

    // // 火曜は「昨日(月曜)が祝日だった場合のみ」送る
    // if (todayDay === 2) {
    //     const yesterday = addDays(today, -1);
    //     if (!isJapaneseHoliday(yesterday, holidaySet)) {
    //         Logger.log('昨日が祝日ではないため、火曜送信はしません。');
    //         return;
    //     }
    // }

    const thisWeekMonday = new Date(today);
    const day = thisWeekMonday.getDay();
    const diff = day === 0 ? -6 : 1 - todayDay;
    thisWeekMonday.setDate(thisWeekMonday.getDate() + diff);

    const weekStart = normalizeDate(thisWeekMonday);
    const weekKey = Utilities.formatDate(weekStart, tz, 'yyyy/MM/dd');
    const label = `*【${weekStart.getMonth() + 1}/${weekStart.getDate()}週】*`;

    const lines = [];
    lines.push(`*【今週の掃除当番】`);
    lines.push('');

    if (skipWeeks.has(weekKey)) {
        lines.push(`${label}\n 長期休暇のため当番なし\n`);
    } else {
        const dutyMembers = getDutyMembersForWeek(
            weekStart, members, dutyCount, baseMonday, baseStartNo, skipWeeks
        )

        const names = dutyMembers.map(m => m.slack_id ? m.slack_id : m.name).join('・');
        lines.push(`${label}\n ${names}`)
    }

    const payload = {
        text: lines.join('\n')
    };

    // Logger.log('▼送信内容▼');
    // Logger.log(payload.text);

    UrlFetchApp.fetch(webhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
    })
}

function getCabinetOfficeHolidaySet() {
    const csvUrl = 'https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv';
    const response = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });

    if (response.getResponseCode() !== 200) {
        throw new Error('内閣府CSVの取得に失敗しました。');
    }

    const text = response.getBlob().getDataAsString('Shift_JIS');
    const csv = Utilities.parseCsv(text);
    const holidaySet = new Set();

    for (let i = 1; i < csv.length; i++) {
        const row = csv[i];
        if (!row || row.length < 1) continue;

        const rawDate = String(row[0] || '').trim();
        if (!rawDate) continue;

        const normalized = normalizeDateString(rawDate);
        if (normalized) {
            holidaySet.add(normalized);
        }
    }

    return holidaySet;
}

function isJapaneseHoliday(date, holidaySet) {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dateStr = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    return holidaySet.has(dateStr);
}

function normalizeDateString(dateStr) {
    const m = String(dateStr).trim().match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (!m) return '';

    const y = m[1];
    const mo = ('0' + m[2]).slice(-2);
    const d = ('0' + m[3]).slice(-2);

    return `${y}/${mo}/${d}`;
}

function onOpen() {
    SpreadsheetApp.getUi().createMenu('掃除当番').addItem('来月のスケジュールを出力', 'outputNextMonthSchedule').addSeparator().addItem('来月のSlack通知を送信', 'sendNextMonthCleaningSummaryToSlack').addToUi();
}