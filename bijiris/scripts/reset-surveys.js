const fs = require("node:fs/promises");
const path = require("node:path");
const { makeDefaultSurveys } = require("../default-surveys");

const ROOT_DIR = path.join(__dirname, "..");
const DATA_DIR = process.env.DATA_DIR || path.join(ROOT_DIR, "data");
const DB_FILE = process.env.DB_FILE || path.join(DATA_DIR, "db.json");

async function readExistingDb() {
  try {
    const raw = await fs.readFile(DB_FILE, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === "ENOENT") {
      return { surveys: [], responses: [] };
    }
    throw error;
  }
}

async function main() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  const db = await readExistingDb();
  const updated = {
    ...db,
    surveys: makeDefaultSurveys(new Date().toISOString()),
    responses: Array.isArray(db.responses) ? db.responses : [],
  };
  await fs.writeFile(DB_FILE, `${JSON.stringify(updated, null, 2)}\n`);
  console.log(`アンケート定義を更新しました: ${DB_FILE}`);
  console.log(`既存回答は ${updated.responses.length} 件保持しました。`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
