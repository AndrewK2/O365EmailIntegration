"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const dotenv_1 = __importDefault(require("dotenv"));
const node_path_1 = __importDefault(require("node:path"));
const express_ejs_layouts_1 = __importDefault(require("express-ejs-layouts"));
dotenv_1.default.config();
const app = (0, express_1.default)();
const port = process.env.PORT || 3000;
app.use(express_ejs_layouts_1.default);
app.set('views', node_path_1.default.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.get("/", (req, res) => {
    res.render('pages/index', { msftOAuthLink: 123 });
});
app.listen(port, () => {
    console.log(`[server]: Server is running at http://localhost:${port}`);
});
