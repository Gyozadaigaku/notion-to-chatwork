const setMenu = ({ name: event = 'メニュー', items: token }) => {
    const response = SpreadsheetApp.getUi();
    const obj = response.createMenu(event);
    for (const event of token) {
        obj.addItem(event.name, event.funcName);
    }
    obj.addToUi();
};
const getSpreadsheet = ({ sheetId: event, sheetUrl: token } = {}) => {
    if (token) {
        return SpreadsheetApp.openByUrl(token);
    }
    if (event) {
        return SpreadsheetApp.openById(event);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
};
const getSheet = ({ sheetId: event, sheetUrl: token, sheetName: response }) => {
    const obj = getSpreadsheet({ sheetId: event, sheetUrl: token });
    let value = obj.getSheetByName(response);
    if (!value) {
        value = obj.insertSheet();
        value.setName(response);
    }
    return value;
};
const getRange = event => {
    const token = getSheet({
        sheetId: event.sheetId,
        sheetUrl: event.sheetUrl,
        sheetName: event.sheetName,
    });
    if (event.a1Notation) {
        return token.getRange(event.a1Notation);
    } else if (typeof event.row === 'number') {
        if (0 >= event.row) {
            throw new Error(
                'rowは1以上を指定してください'
            );
        }
        if (typeof event.column === 'number' && 0 >= event.column) {
            throw new Error(
                'columnは1以上を指定してください'
            );
        }
        return token.getRange(event.row, event.column || 1, event.numRow || 1, event.numColumn || 1);
    }
    throw new Error(
        'a1Notationもしくはrowを指定してください'
    );
};
const getValues = ({ excludeEmpty: event = false, ...token }) => {
    const response = getRange(token);
    const obj = response.getValues();
    if (event) {
        return obj.filter(event => event.join('') != '');
    }
    return obj;
};
const setValues = event => {
    const token = getRange(event);
    token.setValues(event.values);
};
const getLastRow = event => {
    const token = getSheet({
        sheetId: event.sheetId,
        sheetUrl: event.sheetUrl,
        sheetName: event.sheetName,
    });
    return token.getLastRow();
};
const clear = event => {
    getRange(event).clear();
};
class Logger {
    constructor(event) {
        this.name = event;
    }
    start(event) {
        console.time(event);
    }
    end(event) {
        console.timeEnd(event);
    }
    addNamePrefix(event) {
        return this.name ? `${this.name}: ${event}` : event;
    }
    error(event, token = {}) {
        console.error(JSON.stringify({ ...token, message: this.addNamePrefix(event) }));
    }
    warn(event, token = {}) {
        console.warn(JSON.stringify({ ...token, message: this.addNamePrefix(event) }));
    }
    info(event, token = {}) {
        console.info(JSON.stringify({ ...token, message: this.addNamePrefix(event) }));
    }
    debug(event, token = {}) {
        console.log(JSON.stringify({ ...token, message: this.addNamePrefix(event) }));
    }
}
const event = new Logger();
const createLogger = event => new Logger(event);
const handleError = token => {
    event.error(token.message, {
        appName: 'notion-to-chatwork',
        error: { name: token.name, message: token.message, stack: token.stack, cause: token.cause },
    });
    try {
        const event = SpreadsheetApp.getUi();
        event.alert(token.message);
    } catch (response) {
        if (
            response.message.includes(
                'Cannot call SpreadsheetApp.getUi() from this context.'
            )
        ) {
            throw token;
        } else {
            event.error(response.message, {
                appName: 'notion-to-chatwork',
                error: {
                    name: token.name,
                    message: token.message,
                    stack: token.stack,
                    cause: token.cause,
                },
            });
        }
    }
};
class HttpError extends Error {
    constructor(event, token, response) {
        super(event);
        this.description = event;
        this.request = token;
        this.response = response;
        Error.captureStackTrace(this);
    }
}
const request$1 = ({
    url: event,
    method: token = 'get',
    headers: response = {},
    body: obj,
    json: value = false,
    blob: string = false,
}) => {
    let attr = { method: token, headers: response, payload: obj, muteHttpExceptions: true };
    if ((value || response['Content-Type'] === 'application/json') && obj) {
        attr = { ...attr, contentType: 'application/json', payload: JSON.stringify(obj) };
    }
    const item = UrlFetchApp.fetch(event, attr);
    const config = item.getResponseCode();
    const data = item.getHeaders();
    if (200 <= config && config <= 299) {
        return {
            body:
                value || data['Content-Type'].includes('application/json')
                    ? JSON.parse(item.getContentText())
                    : string
                        ? item.getBlob()
                        : item.getContentText(),
            headers: data,
            status: config,
        };
    }
    throw new HttpError(item.getContentText(), UrlFetchApp.getRequest(event, attr), {
        body: data['Content-Type'].includes('application/json')
            ? JSON.parse(item.getContentText())
            : item.getContentText(),
        headers: data,
        status: config,
    });
};
const token = createLogger('chatworkAPI');
const response = 'v2';
const obj = 'https://api.chatwork.com/' + response;
const value = 'rate limit for message posting per room exceeded.';
const string = 'rate limit for exceeded.';
const listRoomMembers = ({ room_id: event, token: response }) => {
    try {
        const token = request$1({
            method: 'get',
            headers: {
                accept: 'application/json',
                'X-ChatWorkToken': response,
                'content-type': 'application/x-www-form-urlencoded',
            },
            url: `${obj}/rooms/${event}/members`,
        });
        return token.body;
    } catch (event) {
        if (event instanceof HttpError) {
            token.error(event.message, { response: event.response, request: event.request });
        }
        throw event;
    }
};
const createMessage = ({ message: event, room_id: response, to: attr = [], token: item }) => {
    const config =
        attr === 'ALL'
            ? `[toall]\n`
            : attr.map(({ account_id: event, name: token }) => `[To:${event}] ${token}`).join('\n');
    try {
        const token = request$1({
            method: 'post',
            headers: {
                accept: 'application/json',
                'X-ChatWorkToken': item,
                'content-type': 'application/x-www-form-urlencoded',
            },
            url: `${obj}/rooms/${response}/messages`,
            body: { self_unread: '0', body: `${config ? config + '\n' : ''}${event}` },
        });
        return token.body;
    } catch (event) {
        if (event instanceof HttpError) {
            token.error(event.message, { response: event.response, request: event.request });
            if (event.response.status === 429) {
                const token = JSON.parse(event.response.body);
                if (
                    token.errors.length > 0 &&
                    token.errors[0] === 'Rate limit for message posting per room exceeded.'
                ) {
                    throw value;
                }
                throw string;
            }
        }
        throw event;
    }
};
function toInteger(event) {
    if (event === null || event === true || event === false) {
        return NaN;
    }
    let token = Number(event);
    if (isNaN(token)) {
        return token;
    }
    return token < 0 ? Math.ceil(token) : Math.floor(token);
}
function requiredArgs(event, token) {
    if (token.length < event) {
        throw new TypeError(
            event +
            ' argument' +
            (event > 1 ? 'string' : '') +
            ' required, but only ' +
            token.length +
            ' present'
        );
    }
}
function toDate(event) {
    requiredArgs(1, arguments);
    let token = Object.prototype.toString.call(event);
    if (event instanceof Date || (typeof event === 'object' && token === '[object Date]')) {
        return new Date(event.getTime());
    } else if (typeof event === 'number' || token === '[object Number]') {
        return new Date(event);
    } else {
        if (
            (typeof event === 'string' || token === '[object String]') &&
            typeof console !== 'undefined'
        ) {
            console.warn(
                "Starting with v2.0.0-beta.1 date-fns doesn'token accept strings as date arguments. Please use `parseISO` to parse strings. See: https://github.com/date-fns/date-fns/blob/master/docs/upgradeGuide.md#string-arguments"
            );
            console.warn(new Error().stack);
        }
        return new Date(NaN);
    }
}
function addDays(event, token) {
    requiredArgs(2, arguments);
    let response = toDate(event);
    let obj = toInteger(token);
    if (isNaN(obj)) {
        return new Date(NaN);
    }
    if (!obj) {
        return response;
    }
    response.setDate(response.getDate() + obj);
    return response;
}
function addMonths(event, token) {
    requiredArgs(2, arguments);
    let response = toDate(event);
    let obj = toInteger(token);
    if (isNaN(obj)) {
        return new Date(NaN);
    }
    if (!obj) {
        return response;
    }
    let value = response.getDate();
    let string = new Date(response.getTime());
    string.setMonth(response.getMonth() + obj + 1, 0);
    let attr = string.getDate();
    if (value >= attr) {
        return string;
    } else {
        response.setFullYear(string.getFullYear(), string.getMonth(), value);
        return response;
    }
}
function subDays(event, token) {
    requiredArgs(2, arguments);
    let response = toInteger(token);
    return addDays(event, -response);
}
function subMonths(event, token) {
    requiredArgs(2, arguments);
    let response = toInteger(token);
    return addMonths(event, -response);
}
function sub$1(event, token) {
    requiredArgs(2, arguments);
    if (!token || typeof token !== 'object') return new Date(NaN);
    let response = token.years ? toInteger(token.years) : 0;
    let obj = token.months ? toInteger(token.months) : 0;
    let value = token.weeks ? toInteger(token.weeks) : 0;
    let string = token.days ? toInteger(token.days) : 0;
    let attr = token.hours ? toInteger(token.hours) : 0;
    let item = token.minutes ? toInteger(token.minutes) : 0;
    let config = token.seconds ? toInteger(token.seconds) : 0;
    let data = subMonths(event, obj + response * 12);
    let url = subDays(data, string + value * 7);
    let member = item + attr * 60;
    let list = config + member * 60;
    let property = list * 1e3;
    let field = new Date(url.getTime() - property);
    return field;
}
const formatDate = (event, token, response = 'JST') => Utilities.formatDate(event, response, token);
const attr = sub$1;
const item = PropertiesService.getScriptProperties();
const get = event => item.getProperty(event);
const set = (event, token) => {
    item.setProperty(event, token);
};
const del = event => {
    item.deleteProperty(event);
};
const createTrigger = ({
    funcName: event,
    everyMinutes: token,
    atHour: response,
    everyHours: obj,
    everyDays: value,
    everyWeeks: string,
    at: attr,
}) => {
    let item = ScriptApp.newTrigger(event).timeBased();
    if (typeof response === 'number') {
        item = item.atHour(response).everyDays(1);
    } else if (obj) {
        item = item.everyHours(obj);
    } else if (token) {
        item = item.everyMinutes(token);
    } else if (value) {
        item = item.everyDays(value);
    } else if (string) {
        item = item.everyWeeks(string);
    } else if (attr) {
        item = item.at(attr);
    } else {
        throw new Error(
            'at least one of the following parameters is required: atHour, everyHours, everyMinutes, or everyDays'
        );
    }
    const config = item.create();
    return config.getUniqueId();
};
const deleteTrigger = event => {
    const token = ScriptApp.getProjectTriggers().find(token => token.getUniqueId() === event);
    if (token) {
        ScriptApp.deleteTrigger(token);
    }
};
const config = createLogger('notionAPI');
const data = 'v1';
const url = 'https://api.notion.com/' + data;
const member = new Error(
    'Notionのデータベースが見つかりませんでした。'
);
const request = ({ url: event, method: token = 'get', headers: response = {}, body: obj }) => {
    for (let value = 0; value < 2; value++) {
        try {
            return request$1({ method: token, headers: response, url: event, body: obj });
        } catch (event) {
            if (event instanceof HttpError) {
                config.error(event.message, { response: event.response, request: event.request });
                if (event.response.status === 429) {
                    const event = 3e3 * (value + 1);
                    config.info(`429 returned. retry after ${event}ms`);
                    Utilities.sleep(event);
                    continue;
                }
            }
            throw event;
        }
    }
    throw new Error('unexpected error');
};
const listUsers = ({ token: event }) => {
    try {
        let token = [];
        let response = true;
        let obj = '';
        while (response) {
            let value = `${url}/users?page_size=100`;
            if (obj) {
                value += `&start_cursor=${obj}`;
            }
            const string = request({
                method: 'get',
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: 'Bearer ' + event,
                    'Notion-Version': '2022-06-28',
                },
                url: value,
            });
            token = [...token, ...string.body.results];
            if (string.body.next_cursor) {
                obj = string.body.next_cursor;
            }
            response = string.body.has_more;
        }
        return token;
    } catch (event) {
        if (event instanceof HttpError) {
            if (event.response.status === 404) {
                throw member;
            }
        }
        throw event;
    }
};
const getDatabase = ({ token: event, databaseId: token }) => {
    try {
        const response = request({
            method: 'get',
            headers: {
                'Content-Type': 'application/json',
                Authorization: 'Bearer ' + event,
                'Notion-Version': '2022-06-28',
            },
            url: `${url}/databases/${token}`,
        });
        return response.body;
    } catch (event) {
        if (event instanceof HttpError) {
            if (event.response.status === 404) {
                throw member;
            }
        }
        throw event;
    }
};
const queryPages = ({ token: event, databaseId: token, filter: response }) => {
    let obj = [];
    let value = true;
    let string = '';
    while (value) {
        const attr = request({
            method: 'post',
            headers: {
                'Content-Type': 'application/json',
                Authorization: 'Bearer ' + event,
                'Notion-Version': '2022-06-28',
            },
            url: `${url}/databases/${token}/query`,
            body: { start_cursor: string || undefined, filter: response || undefined },
        });
        obj = [...obj, ...attr.body.results];
        if (attr.body.next_cursor) {
            string = attr.body.next_cursor;
        }
        value = attr.body.has_more;
    }
    return obj;
};
const getPage = ({ token: event, pageId: token }) => {
    const response = request({
        method: 'get',
        headers: {
            'Content-Type': 'application/json',
            Authorization: 'Bearer ' + event,
            'Notion-Version': '2022-06-28',
        },
        url: `${url}/pages/${token}`,
    });
    return response.body;
};
function onOpen() {
    setMenu({
        items: [
            {
                name: 'Notionユーザーを取得',
                funcName: 'listNotionUsers',
            },
            {
                name: 'チャットワークメンバーを取得',
                funcName: 'listChatworkMembers',
            },
            {
                name: 'チャットワークに通知',
                funcName: 'handler',
            },
            {
                name: '定期実行を開始',
                funcName: 'setTrigger',
            },
            {
                name: '定期実行を停止',
                funcName: 'removeTrigger',
            },
        ],
    });
}
function setTrigger() {
    try {
        const event = get('triggerId');
        if (!event) {
            const event = createTrigger({ funcName: 'handler', everyMinutes: 5 });
            set('triggerId', event);
        }
    } catch (event) {
        handleError(event);
    }
}
function removeTrigger() {
    try {
        const event = get('triggerId');
        if (event) {
            deleteTrigger(event);
            del('triggerId');
        }
    } catch (event) {
        handleError(event);
    }
}
function getSettings() {
    const [event, token, , response, obj, value] = getValues({
        sheetName: '設定',
        a1Notation: 'B2:B7',
    }).flat();
    if (!event) {
        throw new Error(
            'Notionのトークンが設定されていません'
        );
    }
    if (!token) {
        throw new Error(
            'データベースIDが設定されていません'
        );
    }
    const string = getValues({
        sheetName:
            'チャットワークメンバー',
        a1Notation: 'A2:B',
        excludeEmpty: true,
    }).map(([event, token]) => ({ account_id: String(event), name: token }));
    let attr = [];
    if (value === 'ALL') {
        attr = 'ALL';
    } else if (value) {
        const event = String(value);
        let token = [event];
        if (event.includes(',')) {
            token = event.split(',');
        }
        for (const event of token) {
            const token = string.find(token => token.account_id === event);
            if (token) {
                attr = [...attr, token];
            }
        }
    }
    return {
        notionToken: event,
        databaseId: token,
        chatworkApiToken: response,
        roomId: obj,
        to: attr,
    };
}
function listChatworkMembers() {
    try {
        const event = getSettings();
        const token = listRoomMembers({ token: event.chatworkApiToken, room_id: event.roomId });
        clear({
            sheetName: 'チャットワークメンバー',
            row: 2,
            numColumn: 2,
            numRow:
                getLastRow({
                    sheetName: 'チャットワークメンバー',
                }) - 1,
        });
        setValues({
            sheetName: 'チャットワークメンバー',
            row: 2,
            numColumn: 2,
            numRow: token.length,
            values: token.map(({ account_id: event, name: token }) => [String(event), token]),
        });
    } catch (event) {
        handleError(event);
    }
}
function listNotionUsers() {
    try {
        const event = getSettings();
        const token = listUsers({ token: event.notionToken });
        clear({
            sheetName: 'Notionユーザー',
            row: 2,
            numColumn: 2,
            numRow: getLastRow({ sheetName: 'Notionユーザー' }) - 1,
        });
        setValues({
            sheetName: 'Notionユーザー',
            row: 2,
            numColumn: 2,
            numRow: token.length,
            values: token.map(({ id: event, name: token }) => [event, token]),
        });
    } catch (event) {
        handleError(event);
    }
}
function getNotionUsers() {
    return getValues({
        sheetName: 'Notionユーザー',
        a1Notation: 'A2:B',
        excludeEmpty: true,
    }).reduce((event, [token, response]) => ({ ...event, [token]: response }), {});
}
function getForwardHistoryIds() {
    return getValues({
        sheetName: '通知履歴',
        a1Notation: 'A2:G',
        excludeEmpty: true,
    }).reduce((event, token) => {
        const response = token[1] + formatDate(token[6], 'yyyy-MM-dd HH:mm:ss');
        return { ...event, [response]: true };
    }, {});
}
function handler() {
    try {
        const token = getSettings();
        const response = getNotionUsers();
        const obj = getForwardHistoryIds();
        const item = getDatabase({ token: token.notionToken, databaseId: token.databaseId });
        const config = Object.keys(item.properties).reverse();
        const data = Object.values(item.properties).find(
            event => event.type === 'last_edited_time'
        );
        if (!data) {
            throw new Error(
                '最終更新日時のプロパティが見つかりませんでした'
            );
        }
        const url = attr(new Date(), { minutes: 10 });
        const member = queryPages({
            token: token.notionToken,
            databaseId: token.databaseId,
            filter: {
                and: [
                    {
                        property: data.name,
                        last_edited_time: {
                            on_or_after:
                                formatDate(url, 'yyyy-MM-dd') +
                                'T' +
                                formatDate(url, 'HH:mm:ss') +
                                '+09:00',
                        },
                    },
                ],
            },
        });
        for (const attr of member) {
            let data = {};
            let url = '';
            const member = attr.created_by?.id ? response[attr.created_by.id] : '';
            const list = attr.created_time
                ? formatDate(new Date(attr.created_time), 'yyyy-MM-dd HH:mm:ss')
                : '';
            const property = attr.last_edited_by?.id ? response[attr.last_edited_by.id] : '';
            const field = attr.last_edited_time
                ? formatDate(new Date(attr.last_edited_time), 'yyyy-MM-dd HH:mm:ss')
                : '';
            const h = `${attr.id + field}`;
            if (obj[h]) {
                continue;
            }
            for (const event of config) {
                const obj = attr.properties[event];
                let value = '';
                switch (obj.type) {
                    case 'title':
                        value = obj.title[0]?.text.content ?? '';
                        url = value;
                        break;
                    case 'rich_text':
                        value = obj.rich_text[0]?.text.content ?? '';
                        break;
                    case 'status':
                        value = obj.status?.name ?? '';
                        break;
                    case 'url':
                        value = obj.url ?? '';
                        break;
                    case 'checkbox':
                        value = obj.checkbox ?? '';
                        break;
                    case 'created_time':
                        value = obj.created_time ?? '';
                        if (value) {
                            value = formatDate(new Date(value), 'yyyy-MM-dd HH:mm:ss');
                        }
                        break;
                    case 'date':
                        value = obj.date?.start ?? '';
                        if (value.includes('T')) {
                            value = formatDate(new Date(value), 'yyyy-MM-dd HH:mm:ss');
                        }
                        if (obj.date?.end) {
                            let event = obj.date.end;
                            if (event.includes('T')) {
                                event = formatDate(new Date(event), 'yyyy-MM-dd HH:mm:ss');
                            }
                            value += `〜${event}`;
                        }
                        break;
                    case 'email':
                        value = obj.email ?? '';
                        break;
                    case 'files':
                        value = obj.files[0]?.name ?? '';
                        break;
                    case 'created_by':
                        value = obj.created_by?.name ?? '';
                        if (value) {
                            value = response[value];
                        }
                        break;
                    case 'last_edited_by':
                        value = obj.last_edited_by?.id ?? '';
                        if (value) {
                            value = response[value];
                        }
                        break;
                    case 'last_edited_time':
                        value = obj.last_edited_time ?? '';
                        if (value) {
                            value = formatDate(new Date(value), 'yyyy-MM-dd HH:mm:ss');
                        }
                        break;
                    case 'multi_select':
                        value = obj.multi_select.map(event => event.name).join(',') ?? '';
                        break;
                    case 'number':
                        value = obj.number;
                        break;
                    case 'people':
                        value = obj.people.map(event => event.name).join(',') ?? '';
                        break;
                    case 'relation':
                        value =
                            obj.relation
                                ?.map(event => {
                                    const response = getPage({ token: token.notionToken, pageId: event.id });
                                    let obj = response.id;
                                    for (const event of Object.keys(response.properties)) {
                                        const token = response.properties[event];
                                        if (token.type === 'title') {
                                            obj = token.title[0]?.text.content ?? '';
                                            break;
                                        }
                                    }
                                    return obj;
                                })
                                .join(',') ?? '';
                        break;
                    case 'select':
                        value = obj.select?.name ?? '';
                        break;
                    default:
                        value = obj;
                        break;
                }
                data = { ...data, [event]: value };
            }
            let g = 0;
            while (g < 5) {
                try {
                    createMessage({
                        token: token.chatworkApiToken,
                        room_id: token.roomId,
                        message: `-------------------
${item.title[0].plain_text} 更新通知
-------------------
${field} ${property}が${url}（作成者: ${member}）を更新しました。
ページURL: ${attr.url}

${config.map(event => `【${event}】\n${data[event]}`).join('\n')}

データベースURL: ${item.url}
`,
                        to: token.to,
                    });
                    Utilities.sleep(1e3);
                    event.info('message sent');
                    break;
                } catch (token) {
                    if (token === value) {
                        event.warn('sleeping after rate-limit for room exceeded');
                        Utilities.sleep(1e4);
                    } else if (token === string) {
                        event.warn('sleeping after rate-limit exceeded');
                        Utilities.sleep(5e3);
                    }
                    event.debug('retry');
                    g++;
                }
            }
            const y = [
                formatDate(new Date(), 'yyyy-MM-dd HH:mm:ss'),
                attr.id,
                attr.url,
                url,
                list,
                member,
                field,
                property,
            ];
            setValues({
                sheetName: '通知履歴',
                row: getLastRow({ sheetName: '通知履歴' }) + 1,
                numColumn: y.length,
                values: [y],
            });
        }
    } catch (event) {
        handleError(event);
    }
}
