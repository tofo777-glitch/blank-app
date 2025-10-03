
import os, sqlite3, uuid
from pathlib import Path
import pandas as pd
import streamlit as st

APP_TITLE = "Pharmacy Materials Management"
DB_NAME_DEFAULT = "materials_manager.db"
LOGO_PATH = str(Path(__file__).parent / "assets" / "mngha_logo.png")
FOOTER_TEXT = "Developed by: Nawaf Ali Alzaidi"
OVERDUE_DAYS = 14

STATUSES = [
    ("saved","Saved to Complete Later"),
    ("new","New"),
    ("in_process","In Process"),
    ("closed","Closed"),
    ("rejected","Rejected"),
]
DEPTS = (
    "Automation-unit dose pharmacy",
    "Narcotic Pharmacy",
    "in-patient pharmacy",
    "IV Room",
    "Outpatient pharmacy",
    "ER pharmacy",
    "Manufacturing pharmacy",
    "Pharmaceutical Care Administration",
)

def get_db_path() -> Path:
    base = os.environ.get("MM_SHARED_DIR", str(Path(__file__).parent))
    return Path(base) / os.environ.get("MM_DB_NAME", DB_NAME_DEFAULT)

@st.cache_resource(show_spinner=False)
def get_conn() -> sqlite3.Connection:
    p = get_db_path(); p.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(p, check_same_thread=False, isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    init_db(conn); migrate(conn); return conn

def init_db(conn):
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        description TEXT NOT NULL,
        oracle TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        UNIQUE(oracle)
    );
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id TEXT,
        department TEXT,
        item_type TEXT CHECK(item_type IN ('catalog','freetext')),
        material_description TEXT,
        oracle_number TEXT,
        free_text_item TEXT,
        quantity INTEGER NOT NULL DEFAULT 1,
        is_spr INTEGER NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'new',
        timestamp TEXT NOT NULL DEFAULT (datetime('now')),
        status_changed_at TEXT NOT NULL DEFAULT (datetime('now')),
        new_comment_for_manager INTEGER NOT NULL DEFAULT 0,
        new_comment_for_requestor INTEGER NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS comments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        request_id INTEGER NOT NULL REFERENCES requests(id) ON DELETE CASCADE,
        by_role TEXT CHECK(by_role IN ('requestor','manager')),
        by_name TEXT,
        text TEXT NOT NULL,
        at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    """)

def migrate(conn):
    def col_exists(table, col):
        return any(r[1]==col for r in conn.execute(f"PRAGMA table_info({table});"))
    for c, ddl in [
        ("batch_id","ALTER TABLE requests ADD COLUMN batch_id TEXT;"),
        ("status_changed_at","ALTER TABLE requests ADD COLUMN status_changed_at TEXT NOT NULL DEFAULT (datetime('now'));"),
        ("new_comment_for_manager","ALTER TABLE requests ADD COLUMN new_comment_for_manager INTEGER NOT NULL DEFAULT 0;"),
        ("new_comment_for_requestor","ALTER TABLE requests ADD COLUMN new_comment_for_requestor INTEGER NOT NULL DEFAULT 0;"),
    ]:
        if not col_exists("requests", c):
            conn.execute(ddl)

def add_material(conn, desc, oracle):
    conn.execute("INSERT OR IGNORE INTO materials(description, oracle, active) VALUES(?,?,1)", (desc.strip(), oracle.strip()))

def list_materials(conn):
    return conn.execute("SELECT id, description, oracle FROM materials WHERE active=1 ORDER BY description").fetchall()

def submit_rows(conn, dept, items):
    batch = str(uuid.uuid4())
    ids=[]
    for it in items:
        cur = conn.execute(
            """
            INSERT INTO requests(batch_id,department,item_type,material_description,oracle_number,free_text_item,quantity,is_spr,status)
            VALUES (?,?,?,?,?,?,?,?,'new')
            """,
            (batch, dept, it["type"], it.get("material_description"), it.get("oracle_number"),
             it.get("free_text_item"), int(it["quantity"]), it.get("is_spr",0))
        )
        ids.append(cur.lastrowid)
    return batch, ids

def update_status(conn, req_id, status):
    conn.execute("UPDATE requests SET status=?, status_changed_at=datetime('now') WHERE id=?", (status, req_id))

def update_status_batch(conn, batch_id, status):
    conn.execute("UPDATE requests SET status=?, status_changed_at=datetime('now') WHERE batch_id=?", (status, batch_id))

def post_comment(conn, req_id, by_role, text, by_name=None):
    conn.execute("INSERT INTO comments(request_id, by_role, by_name, text) VALUES (?,?,?,?)", (req_id, by_role, by_name, text))
    if by_role=="requestor":
        conn.execute("UPDATE requests SET new_comment_for_manager=1 WHERE id=?", (req_id,))
    else:
        conn.execute("UPDATE requests SET new_comment_for_requestor=1 WHERE id=?", (req_id,))

def mark_comments_read(conn, req_id, for_role):
    if for_role=="manager":
        conn.execute("UPDATE requests SET new_comment_for_manager=0 WHERE id=?", (req_id,))
    else:
        conn.execute("UPDATE requests SET new_comment_for_requestor=0 WHERE id=?", (req_id,))

def contains_multi(text, terms):
    if not terms: return True
    t = (terms or "").lower().split()
    s = (text or "").lower()
    return all(tok in s for tok in t)

def filter_catalog(rows, query):
    q = (query or "").strip()
    if not q: return rows
    return [m for m in rows if contains_multi(m[1], q) or contains_multi(m[2], q)]

def ensure_cart(dept=None):
    if "cart" not in st.session_state: st.session_state.cart = {"department": None, "items": []}
    if dept is not None and st.session_state.cart["department"] not in (None, dept):
        st.warning("Department changed. Clearing current list.")
        st.session_state.cart = {"department": dept, "items": []}
    elif dept is not None and st.session_state.cart["department"] is None:
        st.session_state.cart["department"] = dept

def header_bar():
    c1, c2 = st.columns([1,4])
    with c1:
        try: st.image(LOGO_PATH, width=84)
        except Exception: pass
    with c2:
        st.title(APP_TITLE)
        st.caption("Submit & track materials requests — Materials Management")
    st.divider()

def get_setting(conn, key, default=None):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default
def set_setting(conn, key, value):
    conn.execute("INSERT INTO settings(key, value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))

def counts_by_status(conn, dept):
    data={}
    for k,label in STATUSES:
        c = conn.execute("SELECT COUNT(*) FROM requests WHERE department=? AND status=?", (dept,k)).fetchone()[0]
        od = conn.execute("SELECT COUNT(*) FROM requests WHERE department=? AND status=? AND (julianday('now') - julianday(status_changed_at)) > ?", (dept,k,OVERDUE_DAYS)).fetchone()[0]
        data[k] = {"label":label, "count":c, "overdue":od}
    return data

def page_requestor(conn):
    st.header("New Request")
    dept_options = ["— Select department —"] + list(DEPTS)
    dept_choice = st.selectbox("Department", dept_options, key="req_dept_light")
    dept = None if dept_choice == "— Select department —" else dept_choice
    if not dept:
        st.info("Please select a department to continue."); return
    ensure_cart(dept)

    item_source = st.radio("Item Source", ["From materials list","Not listed (free text)"], horizontal=True, key="item_source_light")
    search_val = ""
    if item_source == "From materials list":
        search_val = st.text_input("Search catalog (Oracle # or Description)", key="search_live_light", placeholder="Type any part… (no Enter needed)")

    with st.form("add_item_form_light", clear_on_submit=False):
        mat_desc = oracle = free_text = None
        qty = 1; is_spr = False
        if item_source == "From materials list":
            mats = list_materials(conn); filtered = filter_catalog(mats, search_val)
            label_map = {m[0]: f"{m[1]} (Oracle: {m[2]})" for m in filtered}
            mat_id = st.selectbox("Select Material", options=[m[0] for m in filtered] if filtered else [], format_func=lambda i: label_map.get(i, ""), key="mat_select_light")
            if filtered and mat_id:
                for m in filtered:
                    if m[0]==mat_id: mat_desc, oracle = m[1], m[2]; break
            st.text_input("Oracle #", value=oracle or "", disabled=True, key="oracle_ro_light")
            st.text_input("Description", value=mat_desc or "", disabled=True, key="desc_ro_light")
            qty = st.number_input("Quantity", min_value=1, step=1, value=1, key="qty_light")
        else:
            free_text = st.text_area("Free‑text Item", placeholder="Describe the item not listed (name, size, any identifiers)", key="free_text_light")
            is_spr = st.checkbox("SPR", key="spr_light")
            qty = st.number_input("Quantity", min_value=1, step=1, value=1, key="qty_ft_light")

        c1,c2 = st.columns([1,1])
        add_clicked = c1.form_submit_button("Add to list")
        submit_clicked = c2.form_submit_button("Submit all items")
        if add_clicked:
            if item_source == "From materials list":
                if not (mat_desc and oracle):
                    st.error("Please select a material before adding.")
                else:
                    st.session_state.cart["items"].append({"type":"catalog","material_description":mat_desc,"oracle_number":oracle,"quantity":int(qty),"is_spr":0})
                    st.success(f"Added: {mat_desc} (Oracle {oracle}) ×{int(qty)}")
            else:
                if not free_text: st.error("Please enter the free‑text item.")
                else:
                    st.session_state.cart["items"].append({"type":"freetext","free_text_item":free_text,"quantity":int(qty),"is_spr":1 if is_spr else 0})
                    st.success("Added free‑text item to list.")
        if submit_clicked:
            items = st.session_state.cart["items"]
            if not items: st.error("Your list is empty.")
            else:
                batch, ids = submit_rows(conn, dept, items)
                st.success(f"Submitted {len(ids)} item(s). Batch: {batch[:8]}…")
                st.session_state.cart = {"department": dept, "items": []}

    st.subheader("Items in your list")
    for i,it in enumerate(st.session_state.cart["items"]):
        with st.container(border=True):
            if it["type"]=="catalog":
                st.write(f"**Catalog** — {it['material_description']} (Oracle: {it['oracle_number']}) × **{it['quantity']}**")
            else:
                st.write(f"**Free‑text** — {it['free_text_item']} × **{it['quantity']}** {'(SPR)' if it.get('is_spr') else ''}")
            if st.button("Remove", key=f"rm_{i}"):
                st.session_state.cart["items"].pop(i); st.rerun()

    st.divider()
    st.header("My Requests — Status")
    counts = counts_by_status(conn, dept)
    cols = st.columns(5)
    for i,(k,label) in enumerate(STATUSES):
        with cols[i]:
            if st.button(f"{label}\n{counts[k]['count']} Records", key=f"statbtn_{k}"):
                st.session_state.req_filter_status = k
            if counts[k]["overdue"] and k not in ("closed","rejected"):
                st.caption(f"{counts[k]['overdue']} Overdue")
    sel = st.session_state.get("req_filter_status","new")
    st.subheader(f"Requests — {dict(STATUSES)[sel]}")
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM requests WHERE department=? AND status=? ORDER BY id DESC", (dept, sel)).fetchall()
    if not rows: st.caption("No requests yet.")
    else:
        for r in rows:
            bell = " 🔔" if r["new_comment_for_requestor"] else ""
            with st.expander(f"#{r['id']} — Qty {r['quantity']} — {r['material_description'] or r['free_text_item']} [{r['status']}] {bell}"):
                st.write(f"**Batch**: {r['batch_id']} • **Submitted**: {r['timestamp']} • **Status since**: {r['status_changed_at']}")
                st.write("**Comments**")
                for c in conn.execute("SELECT * FROM comments WHERE request_id=? ORDER BY id ASC", (r["id"],)):
                    who = "Manager" if c[2]=="manager" else "Requestor"
                    st.markdown(f"- _{who}_ at {c[5]}: {c[4]}")
                ctext = st.text_area(f"Add a comment to #{r['id']}", key=f"req_cmt_{r['id']}")
                colc1,colc2 = st.columns([1,1])
                if colc1.button("Post comment", key=f"req_post_{r['id']}"):
                    if ctext.strip(): post_comment(conn, r["id"], "requestor", ctext.strip()); st.rerun()
                if r["new_comment_for_requestor"]:
                    if colc2.button("Mark comments read", key=f"req_read_{r['id']}"):
                        mark_comments_read(conn, r["id"], "requestor"); st.rerun()

def manager_dashboard(conn):
    new_req = conn.execute("SELECT COUNT(*) FROM requests WHERE status='new'").fetchone()[0]
    new_cmts = conn.execute("SELECT COUNT(*) FROM requests WHERE new_comment_for_manager=1").fetchone()[0]
    ncol = st.columns([1,1,2,2])
    ncol[0].metric("New Requests", new_req)
    ncol[1].metric("New Comments", new_cmts)
    if ncol[2].button("Reload"): st.rerun()
    with ncol[3].popover("Security — Change PIN"):
        new_pin = st.text_input("New PIN", type="password")
        confirm = st.text_input("Confirm PIN", type="password")
        if st.button("Save PIN"):
            if not new_pin: st.error("Enter a PIN.")
            elif new_pin!=confirm: st.error("PINs do not match.")
            else: set_setting(conn, "manager_pin", new_pin); st.success("PIN updated.")

def manager_requests(conn):
    st.subheader("Requests")
    fcols = st.columns(4)
    dept_filter = fcols[0].selectbox("Department", ["All"] + list(DEPTS))
    status_filter = fcols[1].selectbox("Status", ["All"] + [k for k,_ in STATUSES], index=1)
    only_new_comments = fcols[2].checkbox("Only with new comments")
    search = fcols[3].text_input("Search text (Oracle/Description)")
    q = "SELECT * FROM requests WHERE 1=1"; params=[]
    if dept_filter!="All": q += " AND department=?"; params.append(dept_filter)
    if status_filter!="All": q += " AND status=?"; params.append(status_filter)
    if only_new_comments: q += " AND new_comment_for_manager=1"
    conn.row_factory = sqlite3.Row
    rows = conn.execute(q + " ORDER BY status='new' DESC, id DESC", tuple(params)).fetchall()
    # Inline filtering by text
    if search:
        rows = [r for r in rows if contains_multi((r["material_description"] or "") + " " + (r["oracle_number"] or "") + " " + (r["free_text_item"] or ""), search)]
    # Group by batch
    batches={}
    for r in rows: batches.setdefault(r["batch_id"] or f"single-{r['id']}", []).append(r)
    for b_id, items in batches.items():
        with st.expander(f"Batch {b_id[:8]}… — {items[0]['department']} — {len(items)} item(s)"):
            bcols = st.columns(4)
            if bcols[0].button("Mark In Process", key=f"b_inproc_{b_id}"):
                update_status_batch(conn, b_id, "in_process"); st.rerun()
            if bcols[1].button("Close batch", key=f"b_close_{b_id}"):
                update_status_batch(conn, b_id, "closed"); st.rerun()
            if bcols[2].button("Reject batch", key=f"b_rej_{b_id}"):
                update_status_batch(conn, b_id, "rejected"); st.rerun()
            if bcols[3].button("Save for later", key=f"b_save_{b_id}"):
                update_status_batch(conn, b_id, "saved"); st.rerun()
            st.write("---")
            for r in items:
                with st.container(border=True):
                    st.write(f"**#{r['id']}** — Qty {r['quantity']} — {r['material_description'] or r['free_text_item']} — **{r['status']}**")
                    c1,c2,c3 = st.columns([1,1,1])
                    new_status = c1.selectbox("Set status", [s for s,_ in STATUSES], index=[s for s,_ in STATUSES].index(r["status"]) if r["status"] in [s for s,_ in STATUSES] else 1, key=f"sel_{r['id']}")
                    if c2.button("Apply", key=f"apply_{r['id']}"):
                        update_status(conn, r["id"], new_status); st.rerun()
                    if c3.button("Reject item", key=f"rej_{r['id']}"):
                        update_status(conn, r["id"], "rejected"); st.rerun()
                    st.write("**Comments**")
                    for c in conn.execute("SELECT * FROM comments WHERE request_id=? ORDER BY id ASC", (r["id"],)):
                        who = "Manager" if c[2]=="manager" else "Requestor"
                        st.markdown(f"- _{who}_ at {c[5]}: {c[4]}")
                    ctext = st.text_input(f"Reply to #{r['id']}", key=f"mgr_cmt_{r['id']}")
                    cc1,cc2 = st.columns([1,1])
                    if cc1.button("Post reply", key=f"mgr_post_{r['id']}"):
                        if ctext.strip():
                            post_comment(conn, r["id"], "manager", ctext.strip()); mark_comments_read(conn, r["id"], "manager"); st.rerun()
                    if r["new_comment_for_manager"]:
                        if cc2.button("Mark comments reviewed", key=f"mgr_mark_{r['id']}"):
                            mark_comments_read(conn, r["id"], "manager"); st.rerun()

def manager_catalog(conn):
    st.subheader("Materials (list) & Import")
    mats = list_materials(conn)
    st.dataframe([{"ID":m[0],"Description":m[1],"Oracle":m[2]} for m in mats], use_container_width=True, hide_index=True)
    st.write("")
    with st.form("add_mat_light"):
        d = st.text_input("Description")
        o = st.text_input("Oracle #")
        if st.form_submit_button("Add"):
            if d and o: add_material(conn, d, o); st.success("Added"); st.rerun()
            else: st.error("Enter description and Oracle #")
    st.write("")
    st.subheader("Bulk import (Excel/CSV)")
    st.caption("Upload a file with Oracle # and Description columns. Accepts Item#, Oracle, Oracle#, Oracle No + Description/Desc headers.")
    template = pd.DataFrame({"Item#":["122292"], "Description":["5 AMINOLEVULINIC ACID HYDROCHLORIDE 1.5GM POWDER FOR ORAL SOLUTION"]})
    st.download_button("Download CSV template", template.to_csv(index=False).encode("utf-8"), "materials_template.csv", "text/csv")
    up = st.file_uploader("Upload .xlsx or .csv", type=["xlsx","csv"], key="mat_import_light")
    if up is not None:
        try:
            df = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
        except Exception as e:
            st.error(f"Could not read file: {e}"); df=None
        if df is not None:
            mapping = {}
            for c in df.columns:
                cl = str(c).strip().lower()
                if cl in ("item#","item","item no","itemno","oracle","oracle#","oracle no"): mapping[c] = "oracle"
                elif "desc" in cl: mapping[c] = "description"
            df = df.rename(columns=mapping)
            missing = [c for c in ("oracle","description") if c not in df.columns]
            if missing:
                st.error("Missing required column(s): " + ", ".join(missing))
            else:
                df["oracle"] = df["oracle"].astype(str).str.strip().str.replace(r"\.0$","", regex=True)
                df["description"] = df["description"].astype(str).str.strip()
                df = df.dropna(subset=["oracle","description"])
                inserted=skipped=0
                for _,row in df.iterrows():
                    before = conn.total_changes
                    add_material(conn, row["description"], row["oracle"])
                    if conn.total_changes > before: inserted += 1
                    else: skipped += 1
                st.success(f"Import done — added {inserted}, skipped {skipped}.")

def page_manager(conn):
    st.header("Materials Manager")
    if "manager_unlocked_light" not in st.session_state: st.session_state.manager_unlocked_light = False
    pin_cfg = get_setting(conn, "manager_pin", None)
    if not st.session_state.manager_unlocked_light:
        with st.form("pin_form_light"):
            pin = st.text_input(f"Manager PIN {'(default: 1234)' if not pin_cfg else '(PIN is set)'}", type="password")
            ok = st.form_submit_button("Unlock")
        if ok:
            if (pin_cfg is None and pin=="1234") or (pin_cfg is not None and pin==pin_cfg):
                st.session_state.manager_unlocked_light=True; st.success("Unlocked"); st.rerun()
            else: st.error("Invalid PIN")
        st.stop()

    tabs = st.tabs(["Dashboard", "Requests", "Materials Catalog"])
    with tabs[0]: manager_dashboard(conn)
    with tabs[1]: manager_requests(conn)
    with tabs[2]: manager_catalog(conn)

def main():
    conn = get_conn()
    header_bar()
    with st.sidebar:
        role = st.radio("Role", ["Requestor","Materials Manager"], horizontal=True, key="me_role_light")
        st.write("DB path:"); st.code(str(get_db_path()))
    if role=="Requestor": page_requestor(conn)
    else: page_manager(conn)
    st.markdown("<hr/>", unsafe_allow_html=True)
    st.markdown(f"<div style='text-align:center; opacity:0.8; padding:6px;'>{FOOTER_TEXT}</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
