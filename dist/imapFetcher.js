"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ImapFetcher = void 0;
const imap_1 = __importDefault(require("imap"));
class ImapFetcher {
    constructor(emailAddress, accessToken, appConsole) {
        this.emailAddress = emailAddress;
        this.accessToken = accessToken;
        this.appConsole = appConsole;
    }
    Fetch() {
        var _a;
        return __awaiter(this, void 0, void 0, function* () {
            const console = this.appConsole;
            const build_XOAuth2_token = (user, access_token) => Buffer
                .from([`user=${user}`, `auth=Bearer ${access_token}`, '', ''].join('\x01'), 'utf-8')
                .toString('base64');
            const xoauth2 = build_XOAuth2_token(this.emailAddress, this.accessToken);
            console === null || console === void 0 ? void 0 : console.info("Using XOAuth2: " + xoauth2);
            const imap = new imap_1.default({
                password: "",
                user: "",
                host: "outlook.office365.com",
                port: 993,
                tls: true,
                authTimeout: 15000,
                xoauth2: xoauth2,
                debug: (params) => {
                    console === null || console === void 0 ? void 0 : console.log(params);
                }
            });
            (_a = this.appConsole) === null || _a === void 0 ? void 0 : _a.info("Connecting to IMAP...");
            let emailArray = [];
            let flag = false;
            const email = () => __awaiter(this, void 0, void 0, function* () {
                imap.connect();
                return new Promise((resolve, reject) => {
                    imap.once('end', function () {
                        return __awaiter(this, void 0, void 0, function* () {
                            console === null || console === void 0 ? void 0 : console.log('Connection ended');
                            resolve(emailArray);
                            flag = true;
                        });
                    });
                });
            });
            function openInbox(cb) {
                imap.openBox('INBOX', true, cb);
            }
            imap.once('ready', function () {
                try {
                    emailArray = [];
                    flag = false;
                    openInbox(function (err, box) {
                        if (err) {
                            throw new Error("Error on mailbox open: " + err);
                        }
                        imap.search([['ALL'], ['SINCE', '1 Jan, 2024']], function (err, uids) {
                            return __awaiter(this, void 0, void 0, function* () {
                                if (err) {
                                    console === null || console === void 0 ? void 0 : console.log('Search error: ' + err);
                                    imap.end();
                                    return;
                                }
                                console === null || console === void 0 ? void 0 : console.debug("Search complete: ", uids);
                                if (!uids || uids.length == 0) {
                                    imap.end();
                                    return;
                                }
                                const fetch = imap.fetch(uids.sort().reverse().filter((u, i) => i < 10), { bodies: 'HEADER.FIELDS (FROM DATE)' });
                                fetch.on('message', function (msg, seqno) {
                                    console === null || console === void 0 ? void 0 : console.log('Message #%d', seqno);
                                    msg.on('body', function (stream, info) {
                                        let buffer = '';
                                        stream.on('data', function (chunk) {
                                            buffer += chunk.toString('utf8');
                                        });
                                        stream.once('end', function () {
                                            let headers = imap_1.default.parseHeader(buffer);
                                            headers['UID'] = [seqno.toString()];
                                            const emailFields = {};
                                            for (const key in headers) {
                                                emailFields[key] = headers[key][0];
                                            }
                                            emailArray.push(emailFields);
                                        });
                                    });
                                    msg.once('end', function () {
                                        console === null || console === void 0 ? void 0 : console.log(`(#${seqno}) Finished`);
                                    });
                                });
                                fetch.once('error', function (err) {
                                    console === null || console === void 0 ? void 0 : console.log('Fetch error: ' + err);
                                });
                                fetch.once('end', function () {
                                    console === null || console === void 0 ? void 0 : console.log('Done fetching all messages!');
                                    imap.end();
                                });
                            });
                        });
                    });
                }
                catch (e) {
                    console === null || console === void 0 ? void 0 : console.error(e);
                }
            });
            imap.once('error', function (err) {
                console === null || console === void 0 ? void 0 : console.log(err);
            });
            return yield email();
        });
    }
}
exports.ImapFetcher = ImapFetcher;
