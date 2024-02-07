"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const dotenv_1 = __importDefault(require("dotenv"));
const node_path_1 = __importDefault(require("node:path"));
dotenv_1.default.config();
const app = (0, express_1.default)();
const port = process.env.PORT || 3000;
app.use('/css', express_1.default.static(__dirname + '/node_modules/bootstrap/dist/css'));
app.set('view engine', 'ejs');
app.set('views', node_path_1.default.join(__dirname, 'views'));
let ejs = require('ejs');
app.get("/", (req, res) => {
    res.render('pages/index');
});
app.get("/2", (req, res) => {
    let html = ejs.render('Express Server: <b><%= test; %></b>', { test: 123 });
    res.send(html);
});
app.listen(port, () => {
    console.log(`[server]: Server is running at http://localhost:${port}`);
});
