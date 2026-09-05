
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
import sqlite3, io, csv, os
from functools import wraps
from datetime import datetime
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.environ.get("THANN_SECRET_KEY", "change-this-secret-key")
DB = os.path.join(os.path.dirname(__file__), "thann_stock.db")


DATABASE_URL = os.environ.get("DATABASE_URL", "")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

def db():
    if DATABASE_URL:
        import psycopg
        conn = psycopg.connect(DATABASE_URL, row_factory=psycopg.rows.dict_row)
        # small compatibility wrapper: use ? placeholders by translating in execute below
        return PgConn(conn)
    c=sqlite3.connect(DB)
    c.row_factory=sqlite3.Row
    return c

class PgCursor:
    def __init__(self, cur): self.cur=cur
    def execute(self, sql, params=()):
        sql=sql.replace("?", "%s")
        self.cur.execute(sql, params); return self
    def fetchone(self): return self.cur.fetchone()
    def fetchall(self): return self.cur.fetchall()

class PgConn:
    def __init__(self, conn): self.conn=conn
    def cursor(self): return PgCursor(self.conn.cursor())
    def execute(self, sql, params=()):
        cur=PgCursor(self.conn.cursor()); cur.execute(sql,params); return cur
    def executescript(self, sql):
        for statement in sql.split(";"):
            if statement.strip(): self.execute(statement)
    def commit(self): self.conn.commit()
    def close(self): self.conn.close()

def init_db():
    c=db()
    if DATABASE_URL:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
          id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, password TEXT NOT NULL,
          role TEXT NOT NULL DEFAULT 'staff'
        );
        CREATE TABLE IF NOT EXISTS products(
          id SERIAL PRIMARY KEY, category TEXT NOT NULL, code TEXT UNIQUE NOT NULL, name TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS stock(
          product_id INTEGER NOT NULL, branch TEXT NOT NULL, qty INTEGER NOT NULL DEFAULT 0,
          PRIMARY KEY(product_id,branch)
        );
        CREATE TABLE IF NOT EXISTS movements(
          id SERIAL PRIMARY KEY, created_at TEXT NOT NULL, username TEXT NOT NULL,
          branch TEXT NOT NULL, product_id INTEGER NOT NULL, action TEXT NOT NULL,
          qty INTEGER NOT NULL, before_qty INTEGER NOT NULL, after_qty INTEGER NOT NULL
        );
        """)
    else:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'staff');
        CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, category TEXT NOT NULL, code TEXT UNIQUE NOT NULL, name TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS stock(product_id INTEGER NOT NULL, branch TEXT NOT NULL, qty INTEGER NOT NULL DEFAULT 0, PRIMARY KEY(product_id,branch));
        CREATE TABLE IF NOT EXISTS movements(id INTEGER PRIMARY KEY, created_at TEXT NOT NULL, username TEXT NOT NULL, branch TEXT NOT NULL, product_id INTEGER NOT NULL, action TEXT NOT NULL, qty INTEGER NOT NULL, before_qty INTEGER NOT NULL, after_qty INTEGER NOT NULL);
        """)
    # Extra staff identity fields. Safe for existing PostgreSQL/SQLite databases.
    try:
        c.execute("ALTER TABLE users ADD COLUMN employee_name TEXT")
    except Exception:
        pass
    try:
        c.execute("ALTER TABLE users ADD COLUMN employee_id TEXT")
    except Exception:
        pass
    if c.execute("SELECT COUNT(*) AS count FROM users").fetchone()["count"]==0:
        c.execute("INSERT INTO users(username,password,role,employee_name,employee_id) VALUES(?,?,?,?,?)",("admin","thann1234","admin","ผู้ดูแลระบบ","ADMIN"))
    # Keep the old demo staff account usable for compatibility.
    staff_count=c.execute("SELECT COUNT(*) AS count FROM users WHERE role='staff'").fetchone()["count"]
    if staff_count==0:
        c.execute("INSERT INTO users(username,password,role,employee_name,employee_id) VALUES(?,?,?,?,?)",("staff","staff1234","staff","พนักงานตัวอย่าง","STAFF001"))
    if c.execute("SELECT COUNT(*) AS count FROM products").fetchone()["count"] == 0:
        for category,code,name in PRODUCTS:
            c.execute("INSERT INTO products(category,code,name) VALUES(?,?,?)",(category,code,name))
        for p in c.execute("SELECT id FROM products").fetchall():
            for b in BRANCHES: c.execute("INSERT INTO stock(product_id,branch,qty) VALUES(?,?,0)",(p["id"],b))
    c.commit(); c.close()

PRODUCTS = [('Massage Oil', 'EB0409', 'EDEN BREEZE BATH & MASSAGE OIL 295 ML.'), ('Massage Oil', 'AW0415', 'AROMATIC WOOD BATH & MASSAGE OIL 295 ML.'), ('Massage Oil', 'EO0401', 'EASTERN ORCHARD BATH & MASSAGE OIL 295 ML.'), ('Massage Oil', 'EG0401', 'EARL GREY INFUSION BODY & MASSAGE OIL 295 ML.'), ('Massage Oil', 'OE0415', 'ORIENTAL ESSENCE BATH&MASSAGE OIL 295 ML.'), ('Massage Oil', 'MF0406', 'LAVENDER & ROSEMARY BATH & MASSAGE OIL 295 ML.'), ('Shower Gel', 'AW0264', 'AROMATIC WOOD AROMATHERAPY SHOWER GEL 320 ML.'), ('Shower Gel', 'EB0211', 'EDEN BREEZE SHOWER GEL 320ML.'), ('Shower Gel', 'EG0201', 'EARL GREY INFUSION SHOWER GEL 320 ML.'), ('Shower Gel', 'OE0240', 'ORIENTAL ESSENCE AROMATHERAPY SHOWER GEL 320 ML.'), ('Shampoo& Conditioner', 'AW0542', 'AROMATIC WOOD SHAMPOO EXTRA SHINE 250 ML.'), ('Shampoo& Conditioner', 'AW0547', 'AROMATIC WOOD AROMATHERAPY SHAMPOO DETOXIFYING FORMULA 300 ML.'), ('Shampoo& Conditioner', 'EB0510', 'EDEN BREEZE COLOUR TREATED HAIR SHAMPOO 250 ML.'), ('Shampoo& Conditioner', 'OE0532', 'ORIENTAL ESSENCE SHAMPOO DETOX 250 ML.'), ('Shampoo& Conditioner', 'AW0637', 'AROMATIC WOOD AROMATHERAPY CONDITIONER 200 G.'), ('Shampoo& Conditioner', 'EB0608', 'EDEN BREEZE COLOUR TREATED HAIR CONDITIONER 200G.'), ('Shampoo& Conditioner', 'OE0617', 'ORIENTAL ESSENCE AROMATHERAPY CONDITIONER 200 G.'), ('Body Milk', 'AW0345', 'AROMATIC WOOD RICE EXTRACT BODY MILK 320 ML.'), ('Body Milk', 'EO0301', 'EASTERN ORCHARD RICE EXTRACT BODY MILK 320 ML.'), ('Body Milk', 'EG0302', 'EARL GRAY RICE EXTRACT BODY MILK 320 ML.'), ('Body Milk', 'OE0322', 'ORIENTAL ESSENCE BODY MILK 320 ML.'), ('Body Milk', 'RC0311', 'JASMINE BLOSSOM RICE EXTRACT BODY MILK 250 ML.'), ('Facial', 'SC1208', 'SHISO HAIR MASK NANO 100 G.'), ('Facial', 'SC1209', 'SHISO ADVANCE PROTECTIVE HAIR SERUM 100 ML.'), ('Facial', 'SC5156', 'SHISO FACIAL CLEANSER NANO 200 ML.'), ('Facial', 'SC5165', 'PURIFYING FACE WASH 150 G.'), ('Facial', 'SC5155', 'SHISO ASTRINGENT TONER 135 ML.'), ('Facial', 'SC5208', 'HYDRATING EMULSION 100 G.'), ('Facial', 'SC5415', 'THANN REVITALISING COMPLEX 10 G. **แป้งรีไวท์'), ('Facial', 'SC5419', 'THANN REVITALISING ESSENCE 36 ML.. **น้ำรีไวท์'), ('Facial', 'SC5513', 'SHISO FACIAL SERUM 30 ML.'), ('Facial', 'SC5541', 'SHISO FACIAL SUNSCREEN SPF 30 NANO SHISO 40 G.'), ('Facial', 'SC5538', 'SHISO AGE INVERSION FACE CREAM NANO SHISO 40 G.'), ('Facial', 'RC5004', 'RICE LIP BALM 10 G.(PARABEN FREE)'), ('Facial', 'RC5112', 'ASTRINGENT CLEANSING WATER 240 ML.'), ('Facial', 'RC5219', 'RICE EXTRACT MOISTURISING CREAM 80 G.'), ('Facial', 'RC5318', 'RICE OATMEAL FACE SCRUB 480 G.'), ('Facial', 'RC5413', 'DETOXIFYING CLAY MASK 500 G.'), ('อื่นๆ', 'FS0008', 'BLACK UNDERWEAR กางเกงใน'), ('อื่นๆ', 'FS0009', 'SHOWER CAP หมวกอาบน้ำ'), ('อื่นๆ', 'FS1803', 'BIO COMB IN ECO WRAPPER. ( THANN logo ) หวี'), ('อื่นๆ', 'LC1102', 'TIME TO REFRESH 15 G.'), ('อื่นๆ', 'SC0318', 'SHISO BODY BUTTER 350 G.')]
BRANCHES = ['Emporium', 'One Bangkok', 'Chatrium', 'Gaysorn']
init_db()

def login_required(f):
    @wraps(f)
    def w(*a,**k):
        if "user" not in session: return redirect(url_for("login"))
        return f(*a,**k)
    return w

@app.route("/login",methods=["GET","POST"])
def login():
    if request.method=="POST":
        login_type=request.form.get("login_type","admin")
        c=db()
        if login_type=="staff":
            employee_name=request.form.get("employee_name","").strip()
            employee_id=request.form.get("employee_id","").strip()
            if not employee_name or not employee_id:
                c.close(); flash("กรุณากรอกชื่อและรหัสพนักงานให้ครบถ้วน")
                return render_template("login.html")
            # Register/update the staff identity so every movement shows the real name and employee ID.
            existing=c.execute("SELECT * FROM users WHERE employee_id=? AND role='staff'",(employee_id,)).fetchone()
            if existing:
                c.execute("UPDATE users SET employee_name=?, username=? WHERE id=?",(employee_name,employee_id,existing["id"]))
                user=c.execute("SELECT * FROM users WHERE id=?",(existing["id"],)).fetchone()
            else:
                c.execute("INSERT INTO users(username,password,role,employee_name,employee_id) VALUES(?,?,?,?,?)",(employee_id,"", "staff",employee_name,employee_id))
                user=c.execute("SELECT * FROM users WHERE employee_id=? AND role='staff'",(employee_id,)).fetchone()
            c.commit(); c.close()
            session["user"]=user["username"]
            session["display_user"]=f"{user['employee_name']} ({user['employee_id']})"
            session["role"]="staff"
            return redirect(url_for("dashboard"))
        u=request.form.get("username","").strip(); p=request.form.get("password","")
        user=c.execute("SELECT * FROM users WHERE username=? AND password=? AND role='admin'",(u,p)).fetchone(); c.close()
        if user:
            session["user"]=user["username"]; session["display_user"]=user.get("employee_name") or user["username"]; session["role"]="admin"
            return redirect(url_for("dashboard"))
        flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")
    return render_template("login.html")

@app.route("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.route("/")
@login_required
def dashboard():
    branch=request.args.get("branch",BRANCHES[0])
    q=request.args.get("q","").strip()
    cat=request.args.get("category","all")
    c=db()
    params=[branch]
    sql="""SELECT p.*,s.qty FROM products p JOIN stock s ON p.id=s.product_id AND s.branch=? WHERE 1=1"""
    if q: sql+=" AND (LOWER(p.code) LIKE ? OR LOWER(p.name) LIKE ?)"; params += [f"%{q.lower()}%",f"%{q.lower()}%"]
    if cat!="all": sql+=" AND p.category=?"; params.append(cat)
    rows=c.execute(sql,params).fetchall()
    cats=[r[0] for r in c.execute("SELECT DISTINCT category FROM products ORDER BY id").fetchall()]
    total=sum(r["qty"] for r in c.execute("SELECT qty FROM stock WHERE branch=?",(branch,)).fetchall())
    low=c.execute("SELECT COUNT(*) FROM stock WHERE branch=? AND qty BETWEEN 1 AND 4",(branch,)).fetchone()[0]
    out=c.execute("SELECT COUNT(*) FROM stock WHERE branch=? AND qty=0",(branch,)).fetchone()[0]
    c.close()
    return render_template("dashboard.html",rows=rows,branches=BRANCHES,branch=branch,cats=cats,category=cat,q=q,total=total,low=low,out=out)

@app.post("/movement")
@login_required
def movement():
    code=request.form.get("code"); branch=request.form.get("branch"); action=request.form.get("action")
    try: qty=int(request.form.get("qty","0"))
    except: qty=0
    if branch not in BRANCHES or action not in ("receive","issue") or qty<=0 or (action=="receive" and session.get("role")!="admin"):
        flash("ข้อมูลรายการไม่ถูกต้อง"); return redirect(url_for("dashboard",branch=branch))
    c=db(); p=c.execute("SELECT * FROM products WHERE code=?",(code,)).fetchone()
    if not p: c.close(); flash("ไม่พบสินค้า"); return redirect(url_for("dashboard",branch=branch))
    s=c.execute("SELECT qty FROM stock WHERE product_id=? AND branch=?",(p["id"],branch)).fetchone()
    before=s["qty"]; after=before+qty if action=="receive" else before-qty
    if after<0:
        c.close(); flash(f"ไม่สามารถเบิก {qty} ชิ้นได้ เพราะคงเหลือ {before} ชิ้น")
        return redirect(url_for("dashboard",branch=branch))
    c.execute("UPDATE stock SET qty=? WHERE product_id=? AND branch=?",(after,p["id"],branch))
    c.execute("""INSERT INTO movements(created_at,username,branch,product_id,action,qty,before_qty,after_qty)
                 VALUES(?,?,?,?,?,?,?,?)""",(datetime.now().strftime("%Y-%m-%d %H:%M:%S"),session.get("display_user",session["user"]),branch,p["id"],action,qty,before,after))
    c.commit();c.close()
    flash(("รับเข้า" if action=="receive" else "เบิก") + f" {qty} ชิ้น สำเร็จ")
    return redirect(url_for("dashboard",branch=branch))

@app.route("/report")
@login_required
def report():
    c=db()
    data=[]
    for b in BRANCHES:
        r=c.execute("""SELECT COALESCE(SUM(qty),0) total,
          SUM(CASE WHEN qty=0 THEN 1 ELSE 0 END) out,
          SUM(CASE WHEN qty BETWEEN 1 AND 4 THEN 1 ELSE 0 END) low
          FROM stock WHERE branch=?""",(b,)).fetchone()
        data.append((b,r))
    table=c.execute("""SELECT p.category,p.code,p.name,
      MAX(CASE WHEN s.branch=? THEN s.qty END) emp,
      MAX(CASE WHEN s.branch=? THEN s.qty END) oneb,
      MAX(CASE WHEN s.branch=? THEN s.qty END) chat,
      MAX(CASE WHEN s.branch=? THEN s.qty END) gay
      FROM products p JOIN stock s ON s.product_id=p.id GROUP BY p.id ORDER BY p.id""",tuple(BRANCHES)).fetchall()
    c.close();return render_template("report.html",data=data,table=table)

@app.route("/history")
@login_required
def history():
    q=request.args.get("q","").strip()
    branch=request.args.get("branch","all")
    action=request.args.get("action","issue")
    date_from=request.args.get("date_from","").strip()
    date_to=request.args.get("date_to","").strip()
    c=db()
    where=[]; params=[]
    if branch!="all" and branch in BRANCHES:
        where.append("m.branch=?"); params.append(branch)
    if action in ("issue","receive"):
        where.append("m.action=?"); params.append(action)
    if q:
        where.append("(LOWER(p.code) LIKE ? OR LOWER(p.name) LIKE ? OR LOWER(m.username) LIKE ?)")
        qq=f"%{q.lower()}%"; params += [qq,qq,qq]
    if date_from:
        where.append("m.created_at >= ?"); params.append(date_from+" 00:00:00")
    if date_to:
        where.append("m.created_at <= ?"); params.append(date_to+" 23:59:59")
    sql="""SELECT m.*,p.code,p.name FROM movements m JOIN products p ON p.id=m.product_id"""
    if where: sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY m.id DESC LIMIT 500"
    rows=c.execute(sql,params).fetchall(); c.close()
    return render_template("history.html",rows=rows,branches=BRANCHES,q=q,branch=branch,action=action,date_from=date_from,date_to=date_to)

@app.post("/history/delete")
@login_required
def delete_history():
    if session.get("role")!="admin":
        flash("เฉพาะ Admin เท่านั้นที่สามารถลบประวัติได้")
        return redirect(url_for("history"))
    c=db(); c.execute("DELETE FROM movements"); c.commit(); c.close()
    flash("ลบประวัติรับ–เบิกทั้งหมดเรียบร้อยแล้ว")
    return redirect(url_for("history"))

@app.route("/export/<kind>")
@login_required
def export(kind):
    c=db()
    wb=Workbook(); ws=wb.active
    if kind=="summary":
        ws.title="Stock Summary"; ws.append(["หมวดหมู่","รหัสสินค้า","รายการ"]+BRANCHES+["รวม"])
        rows=c.execute("SELECT id,category,code,name FROM products ORDER BY id").fetchall()
        for p in rows:
            nums=[c.execute("SELECT qty FROM stock WHERE product_id=? AND branch=?",(p["id"],b)).fetchone()["qty"] for b in BRANCHES]
            ws.append([p["category"],p["code"],p["name"]]+nums+[sum(nums)])
        filename="THANN_Stock_Summary.xlsx"
    else:
        ws.title="Movement History";ws.append(["วันเวลา","ผู้ใช้งาน","สาขา","ประเภท","รหัสสินค้า","รายการ","จำนวน","ก่อนทำรายการ","คงเหลือ"])
        rows=c.execute("""SELECT m.*,p.code,p.name FROM movements m JOIN products p ON p.id=m.product_id ORDER BY m.id DESC""").fetchall()
        for r in rows: ws.append([r["created_at"],r["username"],r["branch"],"รับสินค้า" if r["action"]=="receive" else "เบิกสินค้า",r["code"],r["name"],r["qty"],r["before_qty"],r["after_qty"]])
        filename="THANN_Stock_History.xlsx"
    c.close(); out=io.BytesIO();wb.save(out);out.seek(0)
    return send_file(out,download_name=filename,as_attachment=True,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)),debug=False)
