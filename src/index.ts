import express, { Express, Request, Response } from "express";
import dotenv from "dotenv";
import path from "node:path";
import expressLayouts from 'express-ejs-layouts';

dotenv.config();

const app: Express = express();
const port = process.env.PORT || 3000;

app.use(expressLayouts);
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.get("/", (req: Request, res: Response) => {
  res.render('pages/index', {msftOAuthLink : 123});
});

app.listen(port, () => {
  console.log(`[server]: Server is running at http://localhost:${port}`);
});
