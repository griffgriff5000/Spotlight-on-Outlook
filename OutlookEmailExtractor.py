#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Outlook Email Extractor (Fast + Modern + Scroll + Auto Names)
- Choose ONE base folder (default: Desktop). App auto-names:
    * Excel:      Emails DD-MM-YYYY - DD-MM-YYYY.xlsx
    * Attachments:Attachments DD-MM-YYYY - DD-MM-YYYY
- Names update when dates change.
- Modern theme if sv-ttk is installed; Light/Dark toggle.
- DD-MM-YYYY date pickers if tkcalendar is installed; text boxes otherwise.
- Scrollable content; header & status fixed.
- Search entire account or a specific folder.
- Attachment type filter (hidden unless 'Only with attachments' is selected).
- Auto-enable 'Save attachments' when filtering for attachments.
- Per-email attachment subfolders + Excel links; exclude inline images (toggle).
"""

import sys, os, time, uuid, tempfile, threading, re, hashlib
from datetime import datetime, timedelta, date
from dataclasses import dataclass
from typing import List, Optional, Dict, Any, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

import pythoncom  # COM init in threads

# ---------- Optional: tkcalendar (no auto-install) ----------
TKCAL_AVAILABLE = False
DateEntry = None

def ensure_tkcalendar():
    global TKCAL_AVAILABLE, DateEntry
    try:
        from tkcalendar import DateEntry as _DE
        DateEntry = _DE
        TKCAL_AVAILABLE = True
    except Exception:
        TKCAL_AVAILABLE = False

# ---------- Optional: sv-ttk theme (no auto-install) ----------
SVTTK_AVAILABLE = False

def ensure_svttk():
    global SVTTK_AVAILABLE
    try:
        import sv_ttk  # noqa: F401
        SVTTK_AVAILABLE = True
    except Exception:
        SVTTK_AVAILABLE = False

# Pandas lazy-imported later
try:
    import pandas as pd
except ImportError:
    pd = None


# ----------------- Model -----------------

@dataclass
class FilterOptions:
    store: str
    folder_path: str                 # "" => store root
    start_date: Optional[datetime]
    end_date: Optional[datetime]
    has_attachments: Optional[bool]
    unread_only: Optional[bool]
    include_subfolders: bool
    subject_contains: str
    from_contains: str
    max_items: int
    require_running: bool
    want_body_preview: bool
    want_attachment_names: bool
    resolve_exchange_addresses: bool
    allowed_exts: Optional[List[str]]           # selected types/custom; None = any
    exclude_inline_images: bool
    save_attachments: bool
    attachments_dir: Optional[str]             # computed from base_dir if saving
    apply_type_to_email_selection: bool        # only when has_attachments == True


# ----------------- Helpers -----------------

def safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def to_us_outlook_datetime(dt: datetime) -> str:
    return dt.strftime("%m/%d/%Y %I:%M %p")

def start_of_day(d: date) -> datetime: return datetime(d.year, d.month, d.day, 0, 0, 0)

def end_of_day(d: date) -> datetime:   return datetime(d.year, d.month, d.day, 23, 59, 59)

def log_gui(textbox: ScrolledText, msg: str) -> None:
    def _append():
        try:
            textbox.configure(state="normal")
            textbox.insert("end", msg + "\n")
            textbox.see("end")
            textbox.configure(state="disabled")
        except Exception:
            pass
    try:
        textbox.after(0, _append)
    except Exception:
        pass

def build_restrict(o: FilterOptions) -> Optional[str]:
    clauses = []
    if o.start_date is not None:
        clauses.append(f"[ReceivedTime] >= '{to_us_outlook_datetime(o.start_date)}'")
    if o.end_date is not None:
        clauses.append(f"[ReceivedTime] <= '{to_us_outlook_datetime(o.end_date)}'")
    if o.has_attachments is True:
        clauses.append("[HasAttachment] = True")
    elif o.has_attachments is False:
        clauses.append("[HasAttachment] = False")
    if o.unread_only is True:
        clauses.append("[Unread] = True")
    elif o.unread_only is False:
        clauses.append("[Unread] = False")
    return " AND ".join(clauses) if clauses else None

def normalize_ext_list(raw: str) -> List[str]:
    out = []
    for part in re.split(r"[,\s;]+", (raw or "").strip().lower()):
        if not part: continue
        out.append(part if part.startswith(".") else f".{part}")
    return out

def sanitize_for_fs(s: str, maxlen: int = 80) -> str:
    s = re.sub(r'[<>:"/\\|?*]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.rstrip(". ")
    return (s[:maxlen].rstrip() or "no_subject")

def dedup_path(path: str) -> str:
    if not os.path.exists(path): return path
    base, ext = os.path.splitext(path); n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"

def file_url(path: str) -> str:
    return "file:///" + os.path.abspath(path).replace("\\", "/")

def get_desktop_folder() -> str:
    # Try OneDrive Desktop, then classic Desktop, then cwd
    candidates = []
    up = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    for root in [os.environ.get("OneDriveCommercial"), os.environ.get("OneDriveConsumer"), os.environ.get("OneDrive")]:
        if root:
            candidates.append(os.path.join(root, "Desktop"))
    candidates.append(os.path.join(up, "Desktop"))
    for c in candidates:
        if c and os.path.isdir(c):
            return c
    return os.getcwd()


# ----------------- Outlook -----------------

def connect_outlook(require_running: bool = False, retries: int = 6, delay: float = 0.8):
    from win32com.client import gencache, Dispatch, GetActiveObject
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    if require_running:
        try:
            app = GetActiveObject("Outlook.Application")
        except Exception as e:
            raise RuntimeError("Outlook is not running. Please open Outlook and try again.") from e
    else:
        try:
            app = GetActiveObject("Outlook.Application")
        except Exception:
            try:
                app = gencache.EnsureDispatch("Outlook.Application")
            except Exception:
                app = Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    for _ in range(retries):
        try:
            ns.Logon("", "", False, False)
            break
        except Exception:
            time.sleep(delay)
    return ns


# ----------------- Extract -----------------

def resolve_smtp_address(mail, resolve_exchange: bool) -> str:
    try:
        addr_type = safe_str(getattr(mail, "SenderEmailType", ""))
        if resolve_exchange and addr_type.upper() == "EX":
            sender = getattr(mail, "Sender", None)
            if sender:
                try:
                    ex_user = sender.GetExchangeUser()
                    if ex_user:
                        smtp = safe_str(getattr(ex_user, "PrimarySmtpAddress", ""))
                        if smtp: return smtp
                except Exception: pass
                try:
                    ex_dl = sender.GetExchangeDistributionList()
                    if ex_dl:
                        smtp = safe_str(getattr(ex_dl, "PrimarySmtpAddress", ""))
                        if smtp: return smtp
                except Exception: pass
        return safe_str(getattr(mail, "SenderEmailAddress", ""))
    except Exception:
        return ""

def attachment_inline_or_hidden(att) -> bool:
    try:
        pa = att.PropertyAccessor
        try:
            if bool(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B")):
                return True
        except Exception:
            pass
        try:
            cid = safe_str(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E"))
            if cid:
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False

def enumerate_attachments(mail, allowed_exts: Optional[List[str]], exclude_inline: bool) -> List[Tuple[str, Any, bool]]:
    rows: List[Tuple[str, Any, bool]] = []
    try:
        atts = getattr(mail, "Attachments", None)
        if not atts: return rows
        count = int(getattr(atts, "Count", 0))
        for i in range(1, count + 1):
            try:
                a = atts.Item(i)
                if exclude_inline and attachment_inline_or_hidden(a):
                    continue
                fname = safe_str(getattr(a, "FileName", "")).strip() or "attachment"
                ext = os.path.splitext(fname)[1].lower()
                matches = True if allowed_exts is None else (ext in allowed_exts)
                rows.append((fname, a, matches))
            except Exception:
                pass
    except Exception:
        pass
    return rows

def save_attachments_for_mail(mail, base_dir: str, subject: str, received: Optional[datetime],
                              entryid: str, allowed_exts: Optional[List[str]],
                              exclude_inline: bool):
    ts = (received.strftime("%Y%m%d_%H%M%S") if received else "nodate")
    h8 = hashlib.sha1((entryid or subject or ts).encode("utf-8", errors="ignore")).hexdigest()[:8]
    subfolder_name = f"{ts}_{sanitize_for_fs(subject, 60)}_{h8}"

    email_dir = ""              # create only when needed
    saved_paths: List[str] = []
    saved_names: List[str] = []

    for fname, att, matches in enumerate_attachments(mail, allowed_exts, exclude_inline):
        if allowed_exts is not None and not matches:
            continue
        if not email_dir:
            email_dir = os.path.join(base_dir, subfolder_name)
            os.makedirs(email_dir, exist_ok=True)
        root, ext = os.path.splitext(fname or "attachment")
        if not ext:
            fname = root + ".bin"
        out_path = dedup_path(os.path.join(email_dir, sanitize_for_fs(fname, 120)))
        try:
            att.SaveAsFile(out_path)
            saved_paths.append(out_path)
            saved_names.append(os.path.basename(out_path))
        except Exception as e:
            print(f"Failed to save attachment '{fname}': {e}", file=sys.stderr)

    return email_dir, saved_paths, saved_names


def iter_folders(ns, root_folder, include_subfolders: bool):
    stack = [(root_folder, safe_str(getattr(root_folder, "FolderPath", "")) or safe_str(getattr(root_folder, "Name", "")))]
    while stack:
        folder, path = stack.pop()
        yield folder, path
        if include_subfolders:
            try:
                subs = getattr(folder, "Folders", None)
                if subs:
                    for i in range(1, int(getattr(subs, "Count", 0)) + 1):
                        sub = subs.Item(i)
                        sub_name = safe_str(getattr(sub, "Name", ""))
                        stack.append((sub, f"{path}/{sub_name}" if path else sub_name))
            except Exception:
                pass

def get_folder_by_path(ns, store_name: str, folder_path: str):
    stores = getattr(ns, "Folders", None)
    if not stores: raise RuntimeError("No Outlook stores found.")
    store = None
    for i in range(1, stores.Count + 1):
        root = stores.Item(i)
        if safe_str(getattr(root, "Name", "")).lower() == store_name.lower():
            store = root; break
    if store is None: raise RuntimeError(f"Store '{store_name}' not found.")
    if not folder_path.strip():
        return store
    folder = store
    for part in folder_path.strip().strip("/").split("/"):
        if part: folder = folder.Folders.Item(part)
    return folder

def list_store_names(ns) -> List[str]:
    names, stores = [], getattr(ns, "Folders", None)
    if stores:
        for i in range(1, stores.Count + 1):
            names.append(safe_str(getattr(stores.Item(i), "Name", "")))
    return names

def list_folder_paths(ns, store_name: str) -> List[str]:
    store = get_folder_by_path(ns, store_name, "")
    paths: List[str] = []
    stack: List[Tuple[Any, List[str]]] = [(store, [])]
    while stack:
        folder, parts = stack.pop()
        subs = getattr(folder, "Folders", None)
        if not subs:
            continue
        try:
            count = int(getattr(subs, "Count", 0))
        except Exception:
            count = 0
        for i in range(count, 0, -1):
            try:
                sub = subs.Item(i)
            except Exception:
                continue
            name = safe_str(getattr(sub, "Name", "")).strip()
            if not name:
                continue
            new_parts = parts + [name]
            paths.append("/".join(new_parts))
            stack.append((sub, new_parts))
    paths.sort(key=lambda path: [segment.lower() for segment in path.split("/")])
    return paths


def run_extraction(options: FilterOptions, out_path: str, logbox: ScrolledText, progress: ttk.Progressbar, status: tk.StringVar):
    try:
        ns = connect_outlook(require_running=options.require_running)
    except Exception as e:
        messagebox.showerror("Outlook Error", f"Failed to connect to Outlook:\n{e}"); return

    try:
        status.set("Resolving folder...")
        log_gui(logbox, "Connecting to Outlook...")
        folder = get_folder_by_path(ns, options.store, options.folder_path)
        log_gui(logbox, f"Using folder: {safe_str(getattr(folder, 'FolderPath', '')) or '(root)'}")

        restriction = build_restrict(options)
        log_gui(logbox, f"Restriction: {restriction}" if restriction else "No Restrict (may be slower)")

        all_rows: List[Dict[str, Any]] = []
        attach_rows: List[Dict[str, Any]] = []
        folder_count = 0
        item_count = 0

        progress.config(mode="indeterminate"); progress.start(10)
        status.set("Scanning...")

        column_set = ",".join([
            "EntryID","ConversationID","Subject","SenderName","SenderEmailAddress",
            "To","CC","BCC","ReceivedTime","Categories","Importance","Size","UnRead","HasAttachment"
        ])

        for f, fpath in iter_folders(ns, folder, options.include_subfolders):
            folder_count += 1
            try:
                items = getattr(f, "Items", None)
                if not items: continue
                try: items.Sort("[ReceivedTime]", True)
                except Exception: pass
                try: items.SetColumns(column_set)
                except Exception: pass

                if restriction:
                    try: items = items.Restrict(restriction)
                    except Exception as e:
                        log_gui(logbox, f"Restrict failed on {fpath}: {e}")

                it = items.GetFirst() if hasattr(items, "GetFirst") else None
                while it:
                    if options.max_items and item_count >= options.max_items: break
                    try:
                        try:
                            if int(getattr(it, "Class", 0)) != 43:
                                it = items.GetNext(); continue
                        except Exception:
                            it = items.GetNext(); continue

                        subj = safe_str(getattr(it, "Subject", ""))
                        sndr_nm = safe_str(getattr(it, "SenderName", ""))
                        smtp_quick = safe_str(getattr(it, "SenderEmailAddress", ""))

                        att_count_quick = 0
                        try:
                            att_count_quick = int(getattr(getattr(it, "Attachments", None), "Count", 0))
                        except Exception:
                            pass

                        if options.has_attachments is True and att_count_quick == 0:
                            it = items.GetNext(); continue
                        if options.has_attachments is False and att_count_quick > 0:
                            it = items.GetNext(); continue

                        if options.subject_contains and options.subject_contains.lower() not in subj.lower():
                            it = items.GetNext(); continue
                        if options.from_contains:
                            fc = options.from_contains.lower()
                            if fc not in sndr_nm.lower() and fc not in smtp_quick.lower():
                                it = items.GetNext(); continue

                        try: received = getattr(it, "ReceivedTime", None)
                        except Exception: received = None
                        entryid = safe_str(getattr(it, "EntryID", ""))

                        names_list: List[str] = []
                        saved_names: List[str] = []
                        saved_paths: List[str] = []
                        email_folder: str = ""

                        need_enum = (
                            (options.apply_type_to_email_selection and options.allowed_exts)
                            or options.want_attachment_names
                            or options.exclude_inline_images
                            or (options.save_attachments and att_count_quick > 0)
                        )
                        att_rows: List[Tuple[str, Any, bool]] = []

                        if need_enum and att_count_quick > 0:
                            att_rows = enumerate_attachments(it, options.allowed_exts, options.exclude_inline_images)
                            if options.apply_type_to_email_selection and options.allowed_exts and not any(m for _, _, m in att_rows):
                                it = items.GetNext(); continue
                            if options.want_attachment_names:
                                if options.allowed_exts:
                                    names_list = [fname for fname, _, m in att_rows if m]
                                else:
                                    names_list = [fname for fname, _, _ in att_rows]
                            if options.save_attachments:
                                base_dir = options.attachments_dir
                                os.makedirs(base_dir, exist_ok=True)
                                email_folder, saved_paths, saved_names = save_attachments_for_mail(
                                    it, base_dir, subj, received, entryid,
                                    allowed_exts=options.allowed_exts,
                                    exclude_inline=options.exclude_inline_images
                                )
                                if saved_paths:
                                    for pth, nm in zip(saved_paths, saved_names):
                                        attach_rows.append({
                                            "ReceivedTime": received.strftime("%Y-%m-%d %H:%M:%S") if received else "",
                                            "Subject": subj,
                                            "SenderEmail": smtp_quick,
                                            "AttachmentName": nm,
                                            "AttachmentPath": pth,
                                            "Link": f'=HYPERLINK("{file_url(pth)}","{nm}")',
                                        })
                                    log_gui(logbox, f"Saved {len(saved_paths)} attachment(s) -> {email_folder}")
                                else:
                                    log_gui(logbox, "No matching attachments to save for this email.")

                        row = {
                            "EntryID": entryid,
                            "ConversationID": safe_str(getattr(it, "ConversationID", "")),
                            "FolderPath": fpath,
                            "Subject": subj,
                            "SenderName": sndr_nm,
                            "SenderEmail": resolve_smtp_address(it, options.resolve_exchange_addresses),
                            "To": safe_str(getattr(it, "To", "")),
                            "CC": safe_str(getattr(it, "CC", "")),
                            "BCC": safe_str(getattr(it, "BCC", "")),
                            "ReceivedTime": received.strftime("%Y-%m-%d %H:%M:%S") if received else "",
                            "Categories": safe_str(getattr(it, "Categories", "")),
                            "Importance": safe_str(getattr(it, "Importance", "")),
                            "Size": safe_str(getattr(it, "Size", "")),
                            "Unread": bool(getattr(it, "UnRead", False)),
                            "AttachmentCount": att_count_quick,
                            "HasAttachments": att_count_quick > 0,
                        }

                        if options.want_body_preview:
                            try:
                                body_preview = safe_str(getattr(it, "Body", ""))[:200].replace("\r\n", " ").replace("\n", " ").strip()
                            except Exception:
                                body_preview = ""
                            row["BodyPreview"] = body_preview

                        if options.want_attachment_names and names_list:
                            row["SavedAttachmentNames"] = ", ".join(names_list)

                        if options.save_attachments:
                            row["SavedAttachmentCount"] = len(saved_paths)
                            if saved_paths and email_folder:
                                row["AttachmentsFolder"] = email_folder
                                row["OpenAttachments"] = f'=HYPERLINK("{file_url(email_folder)}","Open Folder")'

                        all_rows.append(row); item_count += 1
                    finally:
                        try: it = items.GetNext()
                        except Exception: it = None

                log_gui(logbox, f"Scanned: {fpath} (found so far: {item_count})")
                if options.max_items and item_count >= options.max_items: break
            except Exception as e:
                log_gui(logbox, f"Error reading folder {fpath}: {e}")

        progress.stop(); progress.config(mode="determinate", value=0)
        log_gui(logbox, f"Total folders scanned: {folder_count}")
        log_gui(logbox, f"Total emails collected: {len(all_rows)}")
        if options.save_attachments:
            log_gui(logbox, f"Total attachments saved: {len(attach_rows)}")

        global pd
        if pd is None:
            try:
                import pandas as _pd; pd = _pd
            except Exception:
                raise RuntimeError("pandas is not installed. Install with: pip install pandas openpyxl")

        if not all_rows:
            status.set("No results"); messagebox.showinfo("No Results", "No emails matched your filters."); return

        status.set("Writing Excel...")
        df = pd.DataFrame(all_rows)
        nice_cols = [
            "ReceivedTime","Subject","SenderName","SenderEmail","To","CC",
            "Categories","Unread",
            "HasAttachments","AttachmentCount",
            "SavedAttachmentCount","SavedAttachmentNames",
            "OpenAttachments","AttachmentsFolder",
            "Size","FolderPath","ConversationID","EntryID","BodyPreview",
        ]
        cols = [c for c in nice_cols if c in df.columns] + [c for c in df.columns if c not in nice_cols]
        df = df.loc[:, cols]

        meta = {
            "Store": options.store,
            "Folder": options.folder_path or "(root)",
            "Include Subfolders": options.include_subfolders,
            "Start": options.start_date.strftime("%d-%m-%Y %H:%M:%S") if options.start_date else "",
            "End": options.end_date.strftime("%d-%m-%Y %H:%M:%S") if options.end_date else "",
            "Has Attachments": options.has_attachments if options.has_attachments is not None else "Any",
            "Unread Only": options.unread_only if options.unread_only is not None else "Any",
            "Subject Contains": options.subject_contains,
            "From Contains": options.from_contains,
            "Max Items": options.max_items,
            "Exported At": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
            "Body Preview Included": options.want_body_preview,
            "Attachment Names Included": options.want_attachment_names,
            "Resolve Exchange Addresses": options.resolve_exchange_addresses,
            "Selected Types": ", ".join(options.allowed_exts or []) if options.allowed_exts else "Any",
            "Exclude Inline Images": options.exclude_inline_images,
            "Save Attachments": options.save_attachments,
            "Attachments Base Folder": options.attachments_dir or "",
            "Apply Type To Email Selection": options.apply_type_to_email_selection,
        }
        meta_df = pd.DataFrame(list(meta.items()), columns=["Filter", "Value"])

        out_dir = os.path.dirname(out_path) or os.getcwd()
        os.makedirs(out_dir, exist_ok=True)
        with pd.ExcelWriter(out_path, engine="openpyxl") as w:
            df.to_excel(w, sheet_name="Emails", index=False)
            meta_df.to_excel(w, sheet_name="Filters", index=False)
            if attach_rows:
                adf = pd.DataFrame(attach_rows)
                order = [c for c in ["ReceivedTime","Subject","SenderEmail","AttachmentName","AttachmentPath","Link"] if c in adf.columns]
                adf = adf.loc[:, order]
                adf.to_excel(w, sheet_name="Attachments", index=False)

        status.set("Done")
        messagebox.showinfo("Done", f"Exported {len(all_rows)} emails to:\n{out_path}")
        log_gui(logbox, f"Saved to: {out_path}")
    except Exception as e:
        progress.stop(); progress.config(mode="determinate", value=0)
        status.set("Error")
        messagebox.showerror("Error", f"Extraction failed:\n{e}")


# ----------------- Scrollable container -----------------

class VScrollFrame(ttk.Frame):
    """A vertical scrollable frame: put your content under self.body."""
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = ttk.Frame(self.canvas)
        self._window_id = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._sync_width)
        self._bind_mousewheel(self.canvas)
        self._bind_mousewheel(self.body)
        try:
            style = ttk.Style()
            bg = style.lookup("TFrame", "background") or self.body.cget("background")
            self.canvas.configure(bg=bg)
        except Exception:
            pass
    def _sync_width(self, event):
        self.canvas.itemconfigure(self._window_id, width=event.width)
    def _on_mousewheel(self, event):
        if sys.platform == "darwin":
            delta = -1 * int(event.delta)
        else:
            delta = -1 * int(event.delta / 120)
        self.canvas.yview_scroll(delta, "units")
    def _bind_mousewheel(self, widget):
        widget.bind("<Enter>", lambda e: widget.bind_all("<MouseWheel>", self._on_mousewheel))
        widget.bind("<Leave>", lambda e: widget.unbind_all("<MouseWheel>"))
        widget.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        widget.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))


# ----------------- GUI -----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Email Extractor")
        self.geometry("1080x920"); self.minsize(900, 720)

        ensure_svttk()
        self._apply_base_style()

        self.status_var = tk.StringVar(value="Ready")
        self.summary_var = tk.StringVar(value="Pick a store to get started.")
        self.filter_badge_var = tk.StringVar(value="Active filters: none")
        self._folder_cache: Dict[str, List[str]] = {}
        self._auto_paths_locked = False  # if user manually edits, we stop auto-overwriting

        # Header
        hdr = ttk.Frame(self, padding=(10, 10, 10, 0))
        hdr.pack(fill="x")
        title = ttk.Frame(hdr)
        title.pack(side="left", fill="x", expand=True)
        ttk.Label(title, text="Outlook Email Extractor", style="Header.TLabel").pack(anchor="w")
        ttk.Label(title, text="Filter Outlook quickly. Export clean results.", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))
        right = ttk.Frame(hdr); right.pack(side="right")
        ttk.Label(right, text="Theme:").pack(side="left", padx=(0,6))
        self.theme_var = tk.StringVar(value="Light")
        self.cmb_theme = ttk.Combobox(right, values=["Light","Dark"], state="readonly", width=7, textvariable=self.theme_var)
        self.cmb_theme.pack(side="left")
        self.cmb_theme.bind("<<ComboboxSelected>>", lambda *_: self._apply_theme())
        self._apply_theme(initial=True)

        summary = ttk.Frame(self, padding=(12, 4, 12, 8))
        summary.pack(fill="x")
        ttk.Label(summary, textvariable=self.summary_var, style="Summary.TLabel", wraplength=840, justify="left").pack(anchor="w")
        ttk.Label(summary, textvariable=self.filter_badge_var, style="Badge.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Separator(self).pack(fill="x")

        # Scrollable content
        scroll = VScrollFrame(self)
        scroll.pack(fill="both", expand=True)
        root = ttk.Frame(scroll.body, padding=(12, 8, 12, 12))
        root.pack(fill="both", expand=True)
        for i in range(2): root.columnconfigure(i, weight=1)

        # 1) Account / Scope
        f1 = ttk.LabelFrame(root, text="1) Outlook Account / Scope", padding=10)
        f1.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        for i in range(3): f1.columnconfigure(i, weight=1)

        ttk.Label(f1, text="Store:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.cmb_store = ttk.Combobox(f1, values=[], state="readonly")
        self.cmb_store.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        self.cmb_store.bind("<<ComboboxSelected>>", lambda *_: self.refresh_folders())
        self.btn_refresh = ttk.Button(f1, text="Connect / Load Accounts", command=self.refresh_stores)
        self.btn_refresh.grid(row=0, column=2, sticky="e", padx=6, pady=4)

        ttk.Label(f1, text="Folder path (pick or type):").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.folder_var = tk.StringVar(value="")
        self.folder_var.trace_add("write", lambda *_: self._update_summary())
        self.cmb_folder = ttk.Combobox(f1, textvariable=self.folder_var, values=[], state="normal")
        self.cmb_folder.grid(row=1, column=1, columnspan=2, sticky="ew", padx=6, pady=4)
        self.cmb_folder.bind("<KeyRelease>", lambda *_: self._update_summary())
        self.cmb_folder.bind("<FocusOut>", lambda *_: self._update_summary())

        self.var_all_folders = tk.BooleanVar(value=True)  # default: whole mailbox
        ttk.Checkbutton(f1, text="Search entire account (all folders)",
                        variable=self.var_all_folders,
                        command=self._update_folder_controls).grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=2)

        self.var_require_running = tk.BooleanVar(value=False)
        ttk.Checkbutton(f1, text="Require Outlook already open (don't launch Outlook)",
                        variable=self.var_require_running).grid(row=3, column=0, columnspan=3, sticky="w", padx=6, pady=(2, 0))

        # 2) Filters
        f2 = ttk.LabelFrame(root, text="2) Filters", padding=10)
        f2.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        for c in range(6): f2.columnconfigure(c, weight=1)

        ttk.Label(f2, text="Start date:").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ensure_tkcalendar()
        if TKCAL_AVAILABLE:
            self.dt_start = DateEntry(f2, date_pattern="dd-mm-yyyy", locale="en_GB")
            self.dt_start.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
            self.dt_start.bind("<<DateEntrySelected>>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))
        else:
            self.dt_start = ttk.Entry(f2); self.dt_start.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
            self.dt_start.bind("<FocusOut>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))
            self.dt_start.bind("<Return>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))

        ttk.Label(f2, text="End date:").grid(row=0, column=3, sticky="w", padx=6, pady=4)
        if TKCAL_AVAILABLE:
            self.dt_end = DateEntry(f2, date_pattern="dd-mm-yyyy", locale="en_GB")
            self.dt_end.grid(row=0, column=4, sticky="ew", padx=6, pady=4)
            self.dt_end.bind("<<DateEntrySelected>>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))
        else:
            self.dt_end = ttk.Entry(f2); self.dt_end.grid(row=0, column=4, sticky="ew", padx=6, pady=4)
            self.dt_end.bind("<FocusOut>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))
            self.dt_end.bind("<Return>", lambda *_: (self._update_default_paths(force=True), self._update_summary()))

        today_local = datetime.now().date()
        if TKCAL_AVAILABLE:
            self.dt_end.set_date(today_local); self.dt_start.set_date(today_local - timedelta(days=30))
        else:
            self.dt_end.insert(0, today_local.strftime("%d-%m-%Y"))
            self.dt_start.insert(0, (today_local - timedelta(days=30)).strftime("%d-%m-%Y"))

        ttk.Label(f2, text="Attachments:").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        self.var_has_att = tk.StringVar(value="any")
        ttk.Radiobutton(f2, text="Any", variable=self.var_has_att, value="any").grid(row=1, column=1, sticky="w", padx=2)
        ttk.Radiobutton(f2, text="Only with attachments", variable=self.var_has_att, value="yes").grid(row=1, column=2, sticky="w", padx=2)
        ttk.Radiobutton(f2, text="Only without attachments", variable=self.var_has_att, value="no").grid(row=1, column=3, sticky="w", padx=2)

        def _att_changed(*_):
            self._update_type_visibility()
            if self.var_has_att.get() == "yes":
                self.var_save_atts.set(True)
            self._update_summary()
        self.var_has_att.trace_add("write", _att_changed)

        ttk.Label(f2, text="Unread:").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        self.var_unread = tk.StringVar(value="any")
        ttk.Radiobutton(f2, text="Any", variable=self.var_unread, value="any", command=self._update_summary).grid(row=2, column=1, sticky="w", padx=2)
        ttk.Radiobutton(f2, text="Only unread", variable=self.var_unread, value="yes", command=self._update_summary).grid(row=2, column=2, sticky="w", padx=2)
        ttk.Radiobutton(f2, text="Only read", variable=self.var_unread, value="no", command=self._update_summary).grid(row=2, column=3, sticky="w", padx=2)

        self.var_subfolders = tk.BooleanVar(value=True)
        self.chk_subfolders = ttk.Checkbutton(f2, text="Include subfolders", variable=self.var_subfolders, command=self._update_summary)
        self.chk_subfolders.grid(row=3, column=0, sticky="w", padx=6, pady=4)

        ttk.Label(f2, text="Max items (0 = unlimited):").grid(row=3, column=2, sticky="e", padx=6, pady=4)
        self.spn_max = tk.Spinbox(f2, from_=0, to=500000, increment=100, width=12, command=self._update_summary)
        self.spn_max.grid(row=3, column=3, sticky="w", padx=6, pady=4); self.spn_max.delete(0, "end"); self.spn_max.insert(0, "5000")
        self.spn_max.bind("<KeyRelease>", lambda *_: self._update_summary())
        self.spn_max.bind("<FocusOut>", lambda *_: self._update_summary())

        ttk.Label(f2, text="Subject contains:").grid(row=4, column=0, sticky="w", padx=6, pady=4)
        self.ent_subj = ttk.Entry(f2); self.ent_subj.grid(row=4, column=1, columnspan=2, sticky="ew", padx=6, pady=4)
        self.ent_subj.bind("<KeyRelease>", lambda *_: self._update_summary())

        ttk.Label(f2, text="From contains (name or email):").grid(row=4, column=3, sticky="w", padx=6, pady=4)
        self.ent_from = ttk.Entry(f2); self.ent_from.grid(row=4, column=4, columnspan=2, sticky="ew", padx=6, pady=4)
        self.ent_from.bind("<KeyRelease>", lambda *_: self._update_summary())

        # --- FIX: define advanced flags so _gather_options can read them ---
        self.var_body_prev = tk.BooleanVar(value=False)   # include 200-char body preview
        self.var_att_names = tk.BooleanVar(value=False)   # include attachment names column
        self.var_resolve   = tk.BooleanVar(value=True)    # resolve Exchange EX -> SMTP

        # Optional UI for those flags
        ttk.Separator(f2, orient="horizontal").grid(row=5, column=0, columnspan=6, sticky="ew", pady=(6,2))
        ttk.Checkbutton(f2, text="Include body preview (first 200 chars)",
                        variable=self.var_body_prev, command=self._update_summary).grid(row=6, column=0, columnspan=3, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(f2, text="Include attachment names",
                        variable=self.var_att_names, command=self._update_summary).grid(row=6, column=3, columnspan=3, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(f2, text="Resolve Exchange addresses to SMTP",
                        variable=self.var_resolve, command=self._update_summary).grid(row=7, column=0, columnspan=3, sticky="w", padx=6, pady=(2,6))

        # 3) Save location (NEW: pick once; auto-name both outputs)
        f_loc = ttk.LabelFrame(root, text="3) Save location (base folder)", padding=10)
        f_loc.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        f_loc.columnconfigure(1, weight=1)
        ttk.Label(f_loc, text="Base folder (default: Desktop):").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.base_dir_var = tk.StringVar(value=get_desktop_folder())
        self.ent_base_dir = ttk.Entry(f_loc, textvariable=self.base_dir_var)
        self.ent_base_dir.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self.ent_base_dir.bind("<FocusOut>", lambda *_: self._update_default_paths(force=True))
        ttk.Button(f_loc, text="Browse...", command=self.browse_base_dir).grid(row=0, column=2, sticky="e", padx=6, pady=6)



        self.lbl_excel_name = ttk.Label(f_loc, text="Excel file name: (auto)")
        self.lbl_attach_name = ttk.Label(f_loc, text="Attachments folder name: (auto)")
        self.lbl_excel_name.grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0,4))
        self.lbl_attach_name.grid(row=2, column=0, columnspan=3, sticky="w", padx=12, pady=(0,6))

        # 4) Attachment type filter (auto-hidden)
        tf = ttk.LabelFrame(root, text="4) Attachment type filter (shown only when 'Only with attachments' is selected)", padding=10)
        self.typeframe = tf
        tf.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,8))
        self.var_type_pdf   = tk.BooleanVar(value=False)
        self.var_type_img   = tk.BooleanVar(value=False)
        self.var_type_xls   = tk.BooleanVar(value=False)
        self.var_type_doc   = tk.BooleanVar(value=False)
        self.var_type_ppt   = tk.BooleanVar(value=False)
        self.var_type_arc   = tk.BooleanVar(value=False)
        self.var_excl_inline= tk.BooleanVar(value=True)
        ttk.Checkbutton(tf, text="PDF",    variable=self.var_type_pdf, command=self._update_summary).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="Images", variable=self.var_type_img, command=self._update_summary).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="Excel",  variable=self.var_type_xls, command=self._update_summary).grid(row=0, column=2, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="Documents", variable=self.var_type_doc, command=self._update_summary).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="PowerPoint", variable=self.var_type_ppt, command=self._update_summary).grid(row=0, column=4, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="Archives", variable=self.var_type_arc, command=self._update_summary).grid(row=0, column=5, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(tf, text="Exclude inline images (signatures)", variable=self.var_excl_inline, command=self._update_summary)\
            .grid(row=1, column=0, columnspan=3, sticky="w", padx=8, pady=4)
        ttk.Label(tf, text="Custom extensions (comma-separated):").grid(row=1, column=3, sticky="e", padx=6, pady=4)
        self.var_custom_ext = tk.StringVar(value="")
        ttk.Entry(tf, textvariable=self.var_custom_ext).grid(row=1, column=4, columnspan=2, sticky="ew", padx=6, pady=4)
        self.var_custom_ext.trace_add("write", lambda *_: self._update_summary())
        for i in range(6): tf.columnconfigure(i, weight=1)

        # 5) Save attachments (uses auto folder under base)
        fa = ttk.LabelFrame(root, text="5) Save attachments", padding=10)
        fa.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        fa.columnconfigure(1, weight=1)
        self.var_save_atts = tk.BooleanVar(value=False)
        ttk.Checkbutton(fa, text="Save attachments (to the auto-named folder below)", variable=self.var_save_atts,
                        command=self._update_summary)\
            .grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=6)
        ttk.Label(fa, text="Attachments folder (auto):").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.ent_attach_dir = ttk.Entry(fa)
        self.ent_attach_dir.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        self.ent_attach_dir.bind("<KeyRelease>", lambda *_: self._update_summary())
        self.ent_attach_dir.bind("<FocusOut>", lambda *_: self._update_summary())
        ttk.Button(fa, text="Override...", command=self.browse_attach_dir).grid(row=1, column=2, sticky="e", padx=6, pady=6)

        # 6) Output (uses auto file under base)
        f3 = ttk.LabelFrame(root, text="6) Output", padding=10)
        f3.grid(row=5, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        f3.columnconfigure(1, weight=1)
        ttk.Label(f3, text="Excel path (auto):").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.ent_out = ttk.Entry(f3); self.ent_out.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self.ent_out.bind("<KeyRelease>", lambda *_: self._update_summary())
        self.ent_out.bind("<FocusOut>", lambda *_: self._update_summary())
        ttk.Button(f3, text="Override...", command=self.browse_out).grid(row=0, column=2, sticky="e", padx=6, pady=6)


        # Actions / Progress
        act = ttk.Frame(root, padding=(0,2))
        act.grid(row=6, column=0, columnspan=2, sticky="ew", padx=6, pady=(2,8))
        act.columnconfigure(0, weight=1); act.columnconfigure(2, weight=1)
        self.btn_preview = ttk.Button(act, text="Preview Count", command=self.preview_count)
        self.btn_export  = ttk.Button(act, text="Export to Excel", command=self.export_excel)
        self.btn_preview.grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.btn_export.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        self.progress = ttk.Progressbar(act, length=260, mode="determinate")
        self.progress.grid(row=0, column=2, sticky="e", padx=4, pady=2)

        # Log
        logf = ttk.LabelFrame(root, text="Log", padding=10)
        logf.grid(row=7, column=0, columnspan=2, sticky="nsew", padx=6, pady=8)
        logf.rowconfigure(0, weight=1); logf.columnconfigure(0, weight=1)
        self.logbox = ScrolledText(logf, height=12, state="disabled")
        self.logbox.grid(row=0, column=0, sticky="nsew")

        # Status
        bar = ttk.Frame(self, padding=(8,4))
        bar.pack(fill="x", side="bottom")
        ttk.Label(bar, text="Status:", style="StatusCaption.TLabel").pack(side="left")
        ttk.Label(bar, textvariable=self.status_var).pack(side="left", padx=(6, 0))

        self._update_default_paths(force=True)
        self._update_type_visibility()
        self._update_summary()

    # ---- UI helpers ----
    def _apply_base_style(self):
        style = ttk.Style()
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Summary.TLabel", font=("Segoe UI", 10))
        style.configure("Badge.TLabel", font=("Segoe UI", 9, "bold"), foreground="#2563eb")
        style.configure("StatusCaption.TLabel", font=("Segoe UI", 10))

    def _apply_theme(self, initial=False):
        if SVTTK_AVAILABLE:
            import sv_ttk
            sv_ttk.set_theme(self.theme_var.get().lower())

    def _update_folder_controls(self):
        state = "disabled" if self.var_all_folders.get() else "normal"
        self.cmb_folder.configure(state=state)
        self._update_summary()

    def _update_default_paths(self, force=False):
        if self._auto_paths_locked and not force:
            return
        try:
            start_txt = self.dt_start.get()
            end_txt   = self.dt_end.get()
            rng = f"{start_txt} - {end_txt}"
        except Exception:
            rng = datetime.now().strftime("%d-%m-%Y")
        excel_name = f"Emails {rng}.xlsx"
        att_name   = f"Attachments {rng}"
        base = self.base_dir_var.get() or get_desktop_folder()
        self.ent_out.delete(0, "end")
        self.ent_out.insert(0, os.path.join(base, excel_name))
        self.ent_attach_dir.delete(0, "end")
        self.ent_attach_dir.insert(0, os.path.join(base, att_name))
        self.lbl_excel_name.config(text=f"Excel file name: {excel_name}")
        self.lbl_attach_name.config(text=f"Attachments folder name: {att_name}")
        self._update_summary()

    def _update_type_visibility(self):
        show = self.var_has_att.get() == "yes"
        self.typeframe.grid() if show else self.typeframe.grid_remove()
        self._update_summary()

    def _get_date_text(self, widget) -> str:
        try:
            return widget.get().strip()
        except Exception:
            return ""

    def _update_summary(self) -> None:
        store = self.cmb_store.get().strip() or "No mailbox selected"
        scope = "Entire account" if self.var_all_folders.get() else (self.folder_var.get().strip() or "No folder selected")
        start_txt = self._get_date_text(self.dt_start)
        end_txt = self._get_date_text(self.dt_end)
        if start_txt and end_txt:
            date_txt = f"{start_txt} -> {end_txt}"
        elif start_txt:
            date_txt = f"From {start_txt}"
        elif end_txt:
            date_txt = f"Until {end_txt}"
        else:
            date_txt = "Any date"
        out_path = self.ent_out.get().strip() if hasattr(self, "ent_out") else ""
        out_text = os.path.basename(out_path) if out_path else "Auto when exporting"
        self.summary_var.set(f"Mailbox: {store} | Scope: {scope} | Date range: {date_txt} | Excel: {out_text}")

        active = []
        ha = self.var_has_att.get()
        if ha == "yes":
            active.append("attachments only")
        elif ha == "no":
            active.append("no attachments")
        unread = self.var_unread.get()
        if unread == "yes":
            active.append("unread only")
        elif unread == "no":
            active.append("read only")
        if not self.var_subfolders.get():
            active.append("top folder only")
        subj = self.ent_subj.get().strip()
        if subj:
            active.append(f"subject: {subj}")
        sender = self.ent_from.get().strip()
        if sender:
            active.append(f"from: {sender}")
        try:
            max_val = int(self.spn_max.get() or "0")
        except Exception:
            max_val = 0
        if max_val:
            active.append(f"max {max_val}")
        if ha == "yes":
            if any(var.get() for var in (self.var_type_pdf, self.var_type_img, self.var_type_xls, self.var_type_doc, self.var_type_ppt, self.var_type_arc)) or self.var_custom_ext.get().strip():
                active.append("attachment types")
            if not self.var_excl_inline.get():
                active.append("include inline images")
        badge = ", ".join(active) if active else "none"
        self.filter_badge_var.set(f"Active filters: {badge}")

    # ---- browsing ----
    def browse_base_dir(self):
        d = filedialog.askdirectory(initialdir=self.base_dir_var.get())
        if d:
            self.base_dir_var.set(d)
            self._update_default_paths(force=True)

    def browse_out(self):
        f = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx")],
            initialfile=os.path.basename(self.ent_out.get()),
            initialdir=os.path.dirname(self.ent_out.get() or get_desktop_folder())
        )
        if f:
            self._auto_paths_locked = True
            self.ent_out.delete(0,"end"); self.ent_out.insert(0,f)
            self._update_summary()

    def browse_attach_dir(self):
        d = filedialog.askdirectory(initialdir=self.ent_attach_dir.get())
        if d:
            self._auto_paths_locked = True
            self.ent_attach_dir.delete(0,"end"); self.ent_attach_dir.insert(0,d)
            self._update_summary()

    # ---- stores / folders ----
    def refresh_stores(self):
        try:
            ns = connect_outlook(require_running=self.var_require_running.get())
            stores = list_store_names(ns)
            self.cmb_store["values"] = stores
            if stores: self.cmb_store.current(0); self.refresh_folders()
            self._update_summary()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load stores: {e}")

    def refresh_folders(self):
        try:
            ns = connect_outlook(require_running=self.var_require_running.get())
            store = self.cmb_store.get()
            if not store: return
            if store not in self._folder_cache:
                self._folder_cache[store] = list_folder_paths(ns, store)
            self.cmb_folder["values"] = self._folder_cache[store]
            self._update_summary()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load folders: {e}")

    # ---- build options ----
    def _gather_options(self) -> FilterOptions:
        # parse dates
        def parse_date(entry):
            val = entry.get().strip()
            if not val: return None
            try:
                return start_of_day(datetime.strptime(val,"%d-%m-%Y").date())
            except Exception:
                return None
        sd = parse_date(self.dt_start)
        ed = parse_date(self.dt_end)
        if ed: ed = end_of_day(ed.date())

        ha_map = {"any":None,"yes":True,"no":False}
        un_map = {"any":None,"yes":True,"no":False}

        exts = []
        if self.var_type_pdf.get(): exts += [".pdf"]
        if self.var_type_img.get(): exts += [".png",".jpg",".jpeg",".gif",".bmp"]
        if self.var_type_xls.get(): exts += [".xls",".xlsx"]
        if self.var_type_doc.get(): exts += [".doc",".docx"]
        if self.var_type_ppt.get(): exts += [".ppt",".pptx"]
        if self.var_type_arc.get(): exts += [".zip",".rar",".7z"]
        exts += normalize_ext_list(self.var_custom_ext.get())
        exts = exts or None

        return FilterOptions(
            store=self.cmb_store.get(),
            folder_path="" if self.var_all_folders.get() else self.folder_var.get(),
            start_date=sd,
            end_date=ed,
            has_attachments=ha_map[self.var_has_att.get()],
            unread_only=un_map[self.var_unread.get()],
            include_subfolders=self.var_subfolders.get(),
            subject_contains=self.ent_subj.get().strip(),
            from_contains=self.ent_from.get().strip(),
            max_items=int(self.spn_max.get() or 0),
            require_running=self.var_require_running.get(),
            want_body_preview=self.var_body_prev.get(),
            want_attachment_names=self.var_att_names.get(),
            resolve_exchange_addresses=self.var_resolve.get(),
            allowed_exts=exts,
            exclude_inline_images=self.var_excl_inline.get(),
            save_attachments=self.var_save_atts.get(),
            attachments_dir=self.ent_attach_dir.get().strip(),
            apply_type_to_email_selection=(self.var_has_att.get()=="yes")
        )

    # ---- actions ----
    def preview_count(self):
        opts = self._gather_options()
        messagebox.showinfo("Preview", "Preview mode only counts will be added later...")

    def export_excel(self):
        opts = self._gather_options()
        out = self.ent_out.get().strip()
        if not out:
            messagebox.showerror("Error","No output path selected"); return
        t = threading.Thread(
            target=run_extraction,
            args=(opts,out,self.logbox,self.progress,self.status_var),
            daemon=True
        )
        t.start()

# ---- main ----
if __name__=="__main__":
    app = App()
    app.mainloop()
