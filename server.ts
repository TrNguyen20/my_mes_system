import express, { Request, Response, NextFunction } from "express";
import session from "express-session";
import path from "path";
import bcrypt from "bcryptjs";
import multer from "multer";
import * as XLSX from "xlsx";

const app = express();
const PORT = 3000;

// Middleware for body parsing
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Session setup
app.use(
  session({
    secret: process.env.SECRET_KEY || "enterprise_mes_master_2025",
    resave: false,
    saveUninitialized: false,
    cookie: { maxAge: 24 * 60 * 60 * 1000 },
  })
);

// View engine setup
app.set("view engine", "ejs");
app.set("views", path.join(process.cwd(), "views"));

// Flash message helper
declare module "express-session" {
  interface SessionData {
    userId?: number;
    messages?: Array<{ category: string; message: string }>;
  }
}

function flash(req: Request, message: string, category = "info") {
  if (!req.session.messages) {
    req.session.messages = [];
  }
  req.session.messages.push({ category, message });
}

function getFlashedMessages(req: Request) {
  const msgs = req.session.messages || [];
  req.session.messages = [];
  return msgs;
}

// Multer in-memory storage for Excel uploads
const upload = multer({ storage: multer.memoryStorage() });

// --- DATA STRUCTURES ---

interface Operation {
  id: number;
  work_order_id: number;
  name: string;
  sequence: number;
  status: "Pending" | "In Progress" | "Completed";
  start_time: Date | null;
  end_time: Date | null;
  get_op_duration?: () => number;
}

interface WorkOrder {
  id: number;
  po_number: string;
  part_name: string;
  quantity: number;
  status: "Pending" | "In Progress" | "Completed";
  current_step_index: number;
  date_created: Date;
  date_started: Date | null;
  date_completed: Date | null;
  operations: Operation[];
  get_duration?: () => number | null;
}

interface User {
  id: number;
  username: string;
  email: string;
  password_hash: string;
  role: "admin" | "manager" | "operator";
  is_active: boolean;
  date_created: Date;
}

interface AuditLog {
  id: number;
  user_id: number;
  username: string;
  action: string;
  target_type: string;
  target_id: number;
  detail: string;
  timestamp: Date;
  user?: { username: string };
}

// In-memory data store
let nextUserId = 4;
let nextOrderId = 12;
let nextOpId = 20;
let nextLogId = 5;

const users: User[] = [
  {
    id: 1,
    username: "admin",
    email: "admin@mes.com",
    password_hash: bcrypt.hashSync("admin123", 10),
    role: "admin",
    is_active: true,
    date_created: new Date(Date.now() - 30 * 86400000),
  },
  {
    id: 2,
    username: "manager",
    email: "manager@mes.com",
    password_hash: bcrypt.hashSync("manager123", 10),
    role: "manager",
    is_active: true,
    date_created: new Date(Date.now() - 15 * 86400000),
  },
  {
    id: 3,
    username: "operator",
    email: "operator@mes.com",
    password_hash: bcrypt.hashSync("operator123", 10),
    role: "operator",
    is_active: true,
    date_created: new Date(Date.now() - 5 * 86400000),
  },
];

function attachOrderHelpers(o: WorkOrder): WorkOrder {
  o.get_duration = function () {
    if (this.date_started && this.date_completed) {
      const diffMs = this.date_completed.getTime() - this.date_started.getTime();
      return Math.round((diffMs / 60000) * 100) / 100;
    }
    return null;
  };

  o.operations.forEach((op) => {
    op.get_op_duration = function () {
      if (this.start_time && this.end_time) {
        const diffMs = this.end_time.getTime() - this.start_time.getTime();
        return Math.round((diffMs / 60000) * 100) / 100;
      }
      return 0;
    };
  });

  return o;
}

const workOrders: WorkOrder[] = [
  attachOrderHelpers({
    id: 1,
    po_number: "PO-003",
    part_name: "592#TIR3",
    quantity: 1,
    status: "Completed",
    current_step_index: 3,
    date_created: new Date(Date.now() - 120 * 60000),
    date_started: new Date(Date.now() - 110 * 60000),
    date_completed: new Date(Date.now() - 45 * 60000),
    operations: [
      { id: 1, work_order_id: 1, name: "Cắt phôi Laser", sequence: 1, status: "Completed", start_time: new Date(Date.now() - 110 * 60000), end_time: new Date(Date.now() - 90 * 60000) },
      { id: 2, work_order_id: 1, name: "Gia công CNC Phay", sequence: 2, status: "Completed", start_time: new Date(Date.now() - 90 * 60000), end_time: new Date(Date.now() - 65 * 60000) },
      { id: 3, work_order_id: 1, name: "Kiểm tra chất lượng CMM", sequence: 3, status: "Completed", start_time: new Date(Date.now() - 65 * 60000), end_time: new Date(Date.now() - 45 * 60000) },
    ],
  }),
  attachOrderHelpers({
    id: 2,
    po_number: "PO-203",
    part_name: "292#TER3",
    quantity: 3,
    status: "Completed",
    current_step_index: 3,
    date_created: new Date(Date.now() - 90 * 60000),
    date_started: new Date(Date.now() - 80 * 60000),
    date_completed: new Date(Date.now() - 20 * 60000),
    operations: [
      { id: 4, work_order_id: 2, name: "Gia công cơ khí", sequence: 1, status: "Completed", start_time: new Date(Date.now() - 80 * 60000), end_time: new Date(Date.now() - 50 * 60000) },
      { id: 5, work_order_id: 2, name: "Nhiệt luyện", sequence: 2, status: "Completed", start_time: new Date(Date.now() - 50 * 60000), end_time: new Date(Date.now() - 30 * 60000) },
      { id: 6, work_order_id: 2, name: "Mạ Crom hoàn thiện", sequence: 3, status: "Completed", start_time: new Date(Date.now() - 30 * 60000), end_time: new Date(Date.now() - 20 * 60000) },
    ],
  }),
  attachOrderHelpers({
    id: 3,
    po_number: "PO-204",
    part_name: "292#TER4",
    quantity: 5,
    status: "In Progress",
    current_step_index: 1,
    date_created: new Date(Date.now() - 40 * 60000),
    date_started: new Date(Date.now() - 25 * 60000),
    date_completed: null,
    operations: [
      { id: 7, work_order_id: 3, name: "Gia công chuẩn", sequence: 1, status: "In Progress", start_time: new Date(Date.now() - 25 * 60000), end_time: null },
      { id: 8, work_order_id: 3, name: "Kiểm tra dung sai", sequence: 2, status: "Pending", start_time: null, end_time: null },
    ],
  }),
  attachOrderHelpers({
    id: 7,
    po_number: "PO-208",
    part_name: "292#TER8",
    quantity: 13,
    status: "Pending",
    current_step_index: 0,
    date_created: new Date(Date.now() - 15 * 60000),
    date_started: null,
    date_completed: null,
    operations: [
      { id: 9, work_order_id: 7, name: "Cắt phôi nguyên liệu", sequence: 1, status: "Pending", start_time: null, end_time: null },
      { id: 10, work_order_id: 7, name: "Gia công Tiện CNC", sequence: 2, status: "Pending", start_time: null, end_time: null },
      { id: 11, work_order_id: 7, name: "Đóng gói bàn giao", sequence: 3, status: "Pending", start_time: null, end_time: null },
    ],
  }),
  attachOrderHelpers({
    id: 11,
    po_number: "PO-212",
    part_name: "292#TER12",
    quantity: 21,
    status: "Pending",
    current_step_index: 0,
    date_created: new Date(Date.now() - 5 * 60000),
    date_started: null,
    date_completed: null,
    operations: [
      { id: 12, work_order_id: 11, name: "Gia công chuẩn", sequence: 1, status: "Pending", start_time: null, end_time: null },
    ],
  }),
];

const auditLogs: AuditLog[] = [
  {
    id: 1,
    user_id: 1,
    username: "admin",
    action: "SYSTEM_BOOT",
    target_type: "System",
    target_id: 0,
    detail: "Khởi động dịch vụ MES Enterprise trên nền tảng Node.js",
    timestamp: new Date(Date.now() - 180 * 60000),
    user: { username: "admin" },
  },
  {
    id: 2,
    user_id: 1,
    username: "admin",
    action: "CREATE_ROUTING",
    target_type: "WorkOrder",
    target_id: 3,
    detail: "Tạo lệnh PO: PO-204",
    timestamp: new Date(Date.now() - 40 * 60000),
    user: { username: "admin" },
  },
  {
    id: 3,
    user_id: 1,
    username: "admin",
    action: "STEP_ADVANCE",
    target_type: "WorkOrder",
    target_id: 3,
    detail: "Bắt đầu bước 1: Gia công chuẩn",
    timestamp: new Date(Date.now() - 25 * 60000),
    user: { username: "admin" },
  },
];

function logAction(req: Request, action: string, target_type: string, target_id: number, detail = "") {
  const currentUser = getCurrentUser(req);
  const log: AuditLog = {
    id: nextLogId++,
    user_id: currentUser ? currentUser.id : 1,
    username: currentUser ? currentUser.username : "admin",
    action,
    target_type,
    target_id,
    detail,
    timestamp: new Date(),
    user: { username: currentUser ? currentUser.username : "admin" },
  };
  auditLogs.unshift(log);
}

function sync_to_google_sheets() {
  console.log(`[MES Sync] Simulated two-way sync with Google Sheets (Total orders: ${workOrders.length})`);
}

// Current User lookup
function getCurrentUser(req: Request): User | null {
  if (req.session && req.session.userId) {
    const user = users.find((u) => u.id === req.session.userId);
    if (user && user.is_active) {
      return user;
    }
  }
  return null;
}

// Authentication Middleware
function loginRequired(req: Request, res: Response, next: NextFunction) {
  const user = getCurrentUser(req);
  if (!user) {
    return res.redirect("/login");
  }
  next();
}

function roleRequired(...roles: string[]) {
  return (req: Request, res: Response, next: NextFunction) => {
    const user = getCurrentUser(req);
    if (!user || !roles.includes(user.role)) {
      flash(req, "Quyền truy cập bị từ chối!", "danger");
      return res.redirect("/");
    }
    next();
  };
}

// --- ROUTES ---

// Health check
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", app: "MES Enterprise Dashboard" });
});

// Login Page
app.get("/login", (req, res) => {
  if (getCurrentUser(req)) {
    return res.redirect("/");
  }
  const messages = getFlashedMessages(req);
  res.render("login", { messages });
});

app.post("/login", (req, res) => {
  const { username, password } = req.body;
  const user = users.find((u) => u.username === (username || "").trim());

  if (user && bcrypt.compareSync(password || "", user.password_hash)) {
    if (!user.is_active) {
      flash(req, "Tài khoản đã bị vô hiệu hóa!", "danger");
      return res.redirect("/login");
    }
    req.session.userId = user.id;
    return res.redirect("/");
  }

  flash(req, "Sai tài khoản hoặc mật khẩu!", "danger");
  res.redirect("/login");
});

// Logout
app.get("/logout", loginRequired, (req, res) => {
  req.session.destroy(() => {
    res.redirect("/login");
  });
});

// Main Dashboard
app.get("/", loginRequired, (req, res) => {
  const currentUser = getCurrentUser(req)!;
  const search = ((req.query.search as string) || "").trim().toLowerCase();

  let orders = [...workOrders];
  if (search) {
    orders = orders.filter(
      (o) =>
        o.po_number.toLowerCase().includes(search) ||
        o.part_name.toLowerCase().includes(search) ||
        o.status.toLowerCase().includes(search) ||
        o.quantity.toString().includes(search)
    );
  }

  // Sort descending by date_created
  orders.sort((a, b) => b.date_created.getTime() - a.date_created.getTime());

  const stats = {
    pending: workOrders.filter((o) => o.status === "Pending").length,
    progress: workOrders.filter((o) => o.status === "In Progress").length,
    completed: workOrders.filter((o) => o.status === "Completed").length,
  };

  const messages = getFlashedMessages(req);

  res.render("index", {
    current_user: currentUser,
    orders,
    stats,
    search_val: search,
    messages,
  });
});

// Create Order
app.post("/add", loginRequired, roleRequired("admin", "manager"), (req, res) => {
  const po = (req.body.po_number || "").trim();
  const partName = (req.body.part_name || "").trim();
  const quantity = parseInt(req.body.quantity, 10) || 1;

  if (workOrders.some((o) => o.po_number.toLowerCase() === po.toLowerCase())) {
    flash(req, `Lỗi: Mã PO ${po} đã tồn tại!`, "danger");
    return res.redirect("/");
  }

  let opsInput = req.body["operations[]"] || req.body.operations;
  let opsList: string[] = [];
  if (Array.isArray(opsInput)) {
    opsList = opsInput.map((s: string) => s.trim()).filter((s: string) => s.length > 0);
  } else if (typeof opsInput === "string" && opsInput.trim()) {
    opsList = [opsInput.trim()];
  }
  if (opsList.length === 0) {
    opsList = ["Sản xuất chung"];
  }

  const newOrder: WorkOrder = attachOrderHelpers({
    id: nextOrderId++,
    po_number: po,
    part_name: partName,
    quantity,
    status: "Pending",
    current_step_index: 0,
    date_created: new Date(),
    date_started: null,
    date_completed: null,
    operations: opsList.map((name, i) => ({
      id: nextOpId++,
      work_order_id: nextOrderId - 1,
      name,
      sequence: i + 1,
      status: "Pending",
      start_time: null,
      end_time: null,
    })),
  });

  workOrders.unshift(newOrder);
  logAction(req, "CREATE_ROUTING", "WorkOrder", newOrder.id, `PO: ${po}`);
  sync_to_google_sheets();
  flash(req, "Đã mở lệnh sản xuất kèm quy trình công nghệ!", "success");
  res.redirect("/");
});

// Edit Order
app.post("/edit/:id", loginRequired, roleRequired("admin", "manager"), (req, res) => {
  const id = parseInt(req.params.id, 10);
  const order = workOrders.find((o) => o.id === id);

  if (!order) {
    flash(req, "Không tìm thấy đơn hàng!", "danger");
    return res.redirect("/");
  }

  const oldPo = order.po_number;
  order.po_number = (req.body.po_number || "").trim();
  order.part_name = (req.body.part_name || "").trim();
  order.quantity = parseInt(req.body.quantity, 10) || order.quantity;
  order.status = req.body.status || order.status;
  order.current_step_index = parseInt(req.body.current_step_index, 10) || order.current_step_index;

  // Align dates if completed
  if (order.status === "Completed" && !order.date_completed) {
    order.date_completed = new Date();
  } else if (order.status !== "Completed") {
    order.date_completed = null;
  }
  if (order.status === "In Progress" && !order.date_started) {
    order.date_started = new Date();
  }

  logAction(req, "EDIT_FULL", "WorkOrder", id, `Edited PO ${oldPo} -> ${order.po_number} (Step ${order.current_step_index})`);
  sync_to_google_sheets();
  flash(req, "Cập nhật toàn diện thông tin & quy trình thành công!", "success");
  res.redirect("/");
});

// Update / Advance Step
app.get("/update/:id", loginRequired, (req, res) => {
  const id = parseInt(req.params.id, 10);
  const order = workOrders.find((o) => o.id === id);

  if (!order) {
    flash(req, "Không tìm thấy đơn hàng!", "danger");
    return res.redirect("/");
  }

  const ops = [...order.operations].sort((a, b) => a.sequence - b.sequence);

  if (order.status === "Pending") {
    order.status = "In Progress";
    order.date_started = new Date();
    order.current_step_index = 1;
    if (ops.length > 0) {
      ops[0].status = "In Progress";
      ops[0].start_time = new Date();
    }
  } else if (order.status === "In Progress") {
    const currentIndex = order.current_step_index;
    if (currentIndex > 0 && currentIndex <= ops.length) {
      const currentOp = ops[currentIndex - 1];
      currentOp.status = "Completed";
      currentOp.end_time = new Date();
    }

    if (order.current_step_index < ops.length) {
      order.current_step_index += 1;
      const nextOp = ops[order.current_step_index - 1];
      nextOp.status = "In Progress";
      nextOp.start_time = new Date();
    } else {
      order.status = "Completed";
      order.date_completed = new Date();
    }
  }

  logAction(req, "STEP_ADVANCE", "WorkOrder", order.id, `Advanced to step ${order.current_step_index}`);
  sync_to_google_sheets();
  res.redirect("/");
});

// Delete Order
app.get("/delete/:id", loginRequired, roleRequired("admin"), (req, res) => {
  const id = parseInt(req.params.id, 10);
  const index = workOrders.findIndex((o) => o.id === id);

  if (index !== -1) {
    const poRef = workOrders[index].po_number;
    workOrders.splice(index, 1);
    logAction(req, "DELETE", "WorkOrder", id, `Deleted PO: ${poRef}`);
    sync_to_google_sheets();
    flash(req, `Đã xóa lệnh ${poRef}`, "warning");
  }

  res.redirect("/");
});

// Import from Google Sheets
app.get("/import-from-sheets", loginRequired, roleRequired("admin", "manager"), (req, res) => {
  // Pull simulated or default sample rows
  const sampleSheetOrders = [
    { po_number: `PO-SHEET-${Math.floor(100 + Math.random() * 900)}`, part_name: "Trục quay gia tốc", quantity: 80 },
    { po_number: `PO-SHEET-${Math.floor(100 + Math.random() * 900)}`, part_name: "Bánh răng hộp số", quantity: 120 },
  ];

  let imported = 0;
  sampleSheetOrders.forEach((row) => {
    if (!workOrders.some((o) => o.po_number === row.po_number)) {
      const newOrder = attachOrderHelpers({
        id: nextOrderId++,
        po_number: row.po_number,
        part_name: row.part_name,
        quantity: row.quantity,
        status: "Pending",
        current_step_index: 0,
        date_created: new Date(),
        date_started: null,
        date_completed: null,
        operations: [
          { id: nextOpId++, work_order_id: nextOrderId - 1, name: "Gia công chuẩn", sequence: 1, status: "Pending", start_time: null, end_time: null },
        ],
      });
      workOrders.unshift(newOrder);
      imported++;
    }
  });

  logAction(req, "SYNC_SHEETS_PULL", "System", 0, `Pulled ${imported} orders`);
  sync_to_google_sheets();
  flash(req, `📊 Đồng bộ Google Sheets thành công: +${imported} đơn mới`, "success");
  res.redirect("/");
});

// Download Excel Template
app.get("/download-excel-template", (req, res) => {
  const wb = XLSX.utils.book_new();
  const sampleData = [
    ["po_number", "part_name", "quantity"],
    ["PO-SAMPLE-20", "Part Name Example A", 150],
    ["PO-SAMPLE-21", "Part Name Example B", 200],
    ["PO-SAMPLE-22", "Part Name Example C", 85],
  ];

  const ws = XLSX.utils.aoa_to_sheet(sampleData);
  XLSX.utils.book_append_sheet(wb, ws, "Import Template");
  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Disposition", 'attachment; filename="MES_Template.xlsx"');
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.send(buffer);
});

// Import Excel File
app.post("/import-excel", loginRequired, roleRequired("admin", "manager"), upload.single("excel_file"), (req, res) => {
  if (!req.file || !req.file.buffer) {
    flash(req, "Vui lòng tải lên tệp Excel hợp lệ!", "danger");
    return res.redirect("/");
  }

  try {
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<any>(sheet, { header: 1 });

    let imported = 0;
    // Skip header row
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const po = (row[0] !== undefined ? String(row[0]) : "").trim();
      const partName = (row[1] !== undefined ? String(row[1]) : "").trim();
      const qty = parseInt(row[2], 10) || 1;

      if (!po || po === "None" || workOrders.some((o) => o.po_number.toLowerCase() === po.toLowerCase())) {
        continue;
      }

      const newOrder = attachOrderHelpers({
        id: nextOrderId++,
        po_number: po,
        part_name: partName || "Gia công linh kiện",
        quantity: qty,
        status: "Pending",
        current_step_index: 0,
        date_created: new Date(),
        date_started: null,
        date_completed: null,
        operations: [
          { id: nextOpId++, work_order_id: nextOrderId - 1, name: "Gia công cơ khí", sequence: 1, status: "Pending", start_time: null, end_time: null },
        ],
      });

      workOrders.unshift(newOrder);
      imported++;
    }

    logAction(req, "IMPORT_EXCEL", "System", 0, `Imported ${imported}`);
    sync_to_google_sheets();
    flash(req, `✅ Đã nhập ${imported} đơn hàng từ Excel`, "success");
  } catch (err: any) {
    flash(req, `Lỗi đọc file Excel: ${err.message}`, "danger");
  }

  res.redirect("/");
});

// Export CSV
app.get("/export-csv", loginRequired, roleRequired("admin", "manager"), (req, res) => {
  let csvContent = "\ufeffPO Number,Tên sản phẩm,Số lượng,Trạng thái,Lead Time (Min)\n";

  workOrders.forEach((o) => {
    const leadTime = o.get_duration ? o.get_duration() : "";
    const cleanPo = `"${(o.po_number || "").replace(/"/g, '""')}"`;
    const cleanPart = `"${(o.part_name || "").replace(/"/g, '""')}"`;
    csvContent += `${cleanPo},${cleanPart},${o.quantity},"${o.status}",${leadTime || ""}\n`;
  });

  res.setHeader("Content-Type", "text/csv; charset=utf-8");
  res.setHeader("Content-Disposition", 'attachment; filename="report_mes.csv"');
  res.send(csvContent);
});

// Admin: User Management
app.get("/admin/users", loginRequired, roleRequired("admin"), (req, res) => {
  const currentUser = getCurrentUser(req)!;
  const messages = getFlashedMessages(req);

  res.render("admin_users", {
    current_user: currentUser,
    users,
    user_count: users.length,
    messages,
  });
});

// Admin: Create User
app.post("/admin/create-user", loginRequired, roleRequired("admin"), (req, res) => {
  const { username, email, role, password, confirm_password } = req.body;

  if (password !== confirm_password) {
    flash(req, "Mật khẩu xác nhận không khớp!", "danger");
    return res.redirect("/admin/users");
  }

  const cleanUsername = (username || "").trim();
  if (users.some((u) => u.username.toLowerCase() === cleanUsername.toLowerCase())) {
    flash(req, `Tên đăng nhập ${cleanUsername} đã tồn tại!`, "danger");
    return res.redirect("/admin/users");
  }

  const newUser: User = {
    id: nextUserId++,
    username: cleanUsername,
    email: (email || "").trim(),
    role: role === "admin" || role === "manager" ? role : "operator",
    password_hash: bcrypt.hashSync(password || "password123", 10),
    is_active: true,
    date_created: new Date(),
  };

  users.push(newUser);
  logAction(req, "CREATE_USER", "User", newUser.id, `Created ${newUser.username}`);
  flash(req, "Tạo người dùng thành công!", "success");
  res.redirect("/admin/users");
});

// Admin: Toggle User Active
app.post("/admin/toggle-user/:id", loginRequired, roleRequired("admin"), (req, res) => {
  const id = parseInt(req.params.id, 10);
  const currentUser = getCurrentUser(req)!;

  if (currentUser.id !== id) {
    const u = users.find((user) => user.id === id);
    if (u) {
      u.is_active = !u.is_active;
      logAction(req, "TOGGLE_USER", "User", u.id, `Toggle active to ${u.is_active}`);
      flash(req, `Đã cập nhật trạng thái người dùng ${u.username}`, "info");
    }
  }

  res.redirect("/admin/users");
});

// Admin: Delete User
app.post("/admin/delete-user/:id", loginRequired, roleRequired("admin"), (req, res) => {
  const id = parseInt(req.params.id, 10);
  const currentUser = getCurrentUser(req)!;

  if (currentUser.id !== id) {
    const index = users.findIndex((user) => user.id === id);
    if (index !== -1) {
      const deletedName = users[index].username;
      users.splice(index, 1);
      logAction(req, "DELETE_USER", "User", id, `Deleted user ${deletedName}`);
      flash(req, `Đã xóa người dùng ${deletedName}`, "warning");
    }
  }

  res.redirect("/admin/users");
});

// Audit Log View
app.get("/audit-log", loginRequired, roleRequired("admin"), (req, res) => {
  const logs = auditLogs.slice(0, 100);
  res.render("audit_log", { logs });
});

// Start server
app.listen(PORT, "0.0.0.0", () => {
  console.log(`MES Enterprise Server running at http://0.0.0.0:${PORT}`);
});
