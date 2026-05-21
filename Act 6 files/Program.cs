// newest
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using OfficeOpenXml;

// ── ENTRY POINT ──────────────────────────────────────────────
static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

public static class Session
{
    public static string? CurrentUser { get; set; }
}

// ── MAIN FORM (Tab Container) ─────────────────────────────────
class MainForm : Form
{
    TabControl tabs;

    public MainForm()
    {
        Text = "School Information System";
        Size = new Size(850, 600);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;

        tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(new LoginTab());
        tabs.TabPages.Add(new PasswordRecoveryTab());
        tabs.TabPages.Add(new AboutTab());
        tabs.TabPages.Add(new UserManagementTab());
        tabs.TabPages.Add(new EnrollmentTab());
        tabs.TabPages.Add(new GradeEncodingTab());
        tabs.TabPages.Add(new InstructorAssignmentTab());
        tabs.TabPages.Add(new ReportModuleTab());
        Controls.Add(tabs);
    }
}

// ── DATABASE CONNECTION (public class) ────────────────────────
public class Database
{
    private const string ConnectionString =
        "Server=127.0.0.1;" +
        "Port=3306;" +
        "Database=school;" +
        "UserId=root;" +
        "Password=;" +
        "Connection Timeout=10;";

    public static MySqlConnection GetConnection()
    {
        var conn = new MySqlConnection(ConnectionString);
        conn.Open();
        return conn;
    }

    public static bool TestConnection()
    {
        try
        {
            using var conn = GetConnection();
            return conn.State == ConnectionState.Open;
        }
        catch { return false; }
    }
}

// ── HELPERS ───────────────────────────────────────────────────
static class UI
{
    public static Label MakeLabel(string text, int x, int y)
        => new Label { Text = text, Location = new Point(x, y), AutoSize = true };

    public static TextBox MakeTextBox(int x, int y, int w = 220, bool password = false)
        => new TextBox
        {
            Location = new Point(x, y),
            Width = w,
            PasswordChar = password ? '*' : '\0'
        };

    public static ComboBox MakeCombo(int x, int y, int w, string[] items)
    {
        var c = new ComboBox
        {
            Location = new Point(x, y),
            Width = w,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        c.Items.AddRange(items);
        return c;
    }

    public static Button MakeButton(string text, int x, int y, int w = 110)
        => new Button { Text = text, Location = new Point(x, y), Width = w };

    public static Label MakeTitle(string text)
        => new Label
        {
            Text = text,
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(12, 12)
        };
}

// ════════════════════════════════════════════════════════════
// 1. LOGIN
// ════════════════════════════════════════════════════════════
class LoginTab : TabPage
{
    TextBox txtUser, txtPass;
    ComboBox cmbRole;
    Label lblMsg;

    public LoginTab()
    {
        Text = "Login";
        Controls.Add(UI.MakeTitle("User Login"));

        Controls.Add(UI.MakeLabel("Username:", 20, 60));
        txtUser = UI.MakeTextBox(120, 57);
        Controls.Add(txtUser);

        Controls.Add(UI.MakeLabel("Password:", 20, 95));
        txtPass = UI.MakeTextBox(120, 92, password: true);
        Controls.Add(txtPass);

        Controls.Add(UI.MakeLabel("Role:", 20, 130));
        cmbRole = UI.MakeCombo(120, 127, 220, new[] { "Administrator", "Staff", "Viewer" });
        Controls.Add(cmbRole);

        var btnLogin = UI.MakeButton("Log In", 120, 170);
        btnLogin.Click += BtnLogin_Click;
        Controls.Add(btnLogin);

        var btnClear = UI.MakeButton("Clear", 240, 170);
        btnClear.Click += (s, e) =>
        {
            txtUser.Clear(); txtPass.Clear();
            cmbRole.SelectedIndex = -1; lblMsg.Text = "";
        };
        Controls.Add(btnClear);

        lblMsg = new Label { Location = new Point(120, 210), AutoSize = true, ForeColor = Color.Red };
        Controls.Add(lblMsg);

        var lnk = new LinkLabel { Text = "Forgot password?", Location = new Point(120, 240), AutoSize = true };
        lnk.LinkClicked += (s, e) => ((TabControl)Parent).SelectedIndex = 1;
        Controls.Add(lnk);
    }

    void BtnLogin_Click(object? sender, EventArgs e)
    {
        if (txtUser.Text.Trim() == "" || txtPass.Text.Trim() == "")
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "Please enter username and password.";
            return;
        }

        try
        {
            using var conn = Database.GetConnection();
            const string sql =
                "SELECT COUNT(*) FROM accounts " +
                "WHERE username=@u AND password_hash=@p AND is_active=1";
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@u", txtUser.Text.Trim());
            cmd.Parameters.AddWithValue("@p", txtPass.Text.Trim());
            long count = (long)cmd.ExecuteScalar()!;

            if (count > 0)
            {
                lblMsg.ForeColor = Color.Green;
                Session.CurrentUser = txtUser.Text.Trim();
                lblMsg.Text = "Login successful! Welcome, " + txtUser.Text.Trim() + ".";
                ((TabControl)Parent).SelectedIndex = 3;
            }
            else
            {
                lblMsg.ForeColor = Color.Red;
                lblMsg.Text = "Invalid credentials.";
            }
        }
        catch (Exception ex)
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "DB error: " + ex.Message;
        }
    }
}

// ════════════════════════════════════════════════════════════
// 2. PASSWORD RECOVERY
// ════════════════════════════════════════════════════════════
class PasswordRecoveryTab : TabPage
{
    TextBox txtUser, txtAnswer, txtNewPass, txtConfirm;
    ComboBox cmbQuestion;
    Label lblMsg;

    public PasswordRecoveryTab()
    {
        Text = "Password Recovery";
        Controls.Add(UI.MakeTitle("Password Recovery"));

        Controls.Add(UI.MakeLabel("Username:", 20, 60));
        txtUser = UI.MakeTextBox(180, 57, 200);
        Controls.Add(txtUser);

        Controls.Add(UI.MakeLabel("Security Question:", 20, 95));
        cmbQuestion = UI.MakeCombo(180, 92, 340, new[]
        {
            "What is your mother's maiden name?",
            "What was the name of your first pet?",
            "What city were you born in?",
            "What is your elementary school name?"
        });
        Controls.Add(cmbQuestion);

        Controls.Add(UI.MakeLabel("Answer:", 20, 130));
        txtAnswer = UI.MakeTextBox(180, 127, 200);
        Controls.Add(txtAnswer);

        Controls.Add(UI.MakeLabel("New Password:", 20, 165));
        txtNewPass = UI.MakeTextBox(180, 162, 200, password: true);
        Controls.Add(txtNewPass);

        Controls.Add(UI.MakeLabel("Confirm Password:", 20, 200));
        txtConfirm = UI.MakeTextBox(180, 197, 200, password: true);
        Controls.Add(txtConfirm);

        var btnSend = UI.MakeButton("Reset Password", 180, 235, 130);
        btnSend.Click += BtnReset_Click;
        Controls.Add(btnSend);

        var btnBack = UI.MakeButton("Back to Login", 320, 235, 120);
        btnBack.Click += (s, e) => { if (Parent is TabControl tc) tc.SelectedIndex = 0; };
        Controls.Add(btnBack);

        lblMsg = new Label { Location = new Point(180, 275), AutoSize = true };
        Controls.Add(lblMsg);
    }

    void BtnReset_Click(object? sender, EventArgs e)
    {
        if (txtUser.Text.Trim() == "" || cmbQuestion.SelectedIndex < 0 ||
            txtAnswer.Text.Trim() == "" || txtNewPass.Text.Trim() == "" ||
            txtConfirm.Text.Trim() == "")
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "Please fill in all fields.";
            return;
        }

        if (txtNewPass.Text != txtConfirm.Text)
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "Passwords do not match.";
            return;
        }

        if (txtNewPass.Text.Length < 6)
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "Password must be at least 6 characters.";
            return;
        }

        try
        {
            using var conn = Database.GetConnection();

            const string check =
                "SELECT COUNT(*) FROM accounts " +
                "WHERE username=@u AND security_question=@q AND security_answer=@a";
            using var cmd = new MySqlCommand(check, conn);
            cmd.Parameters.AddWithValue("@u", txtUser.Text.Trim());
            cmd.Parameters.AddWithValue("@q", cmbQuestion.SelectedItem?.ToString() ?? "");
            cmd.Parameters.AddWithValue("@a", txtAnswer.Text.Trim());
            long found = (long)cmd.ExecuteScalar()!;

            if (found == 0)
            {
                lblMsg.ForeColor = Color.Red;
                lblMsg.Text = "Username or security answer is incorrect.";
                return;
            }

            const string upd = "UPDATE accounts SET password_hash=@p WHERE username=@u";
            using var uCmd = new MySqlCommand(upd, conn);
            uCmd.Parameters.AddWithValue("@p", txtNewPass.Text.Trim());
            uCmd.Parameters.AddWithValue("@u", txtUser.Text.Trim());
            uCmd.ExecuteNonQuery();

            lblMsg.ForeColor = Color.Green;
            lblMsg.Text = "Password reset successfully!";
            txtUser.Clear(); txtAnswer.Clear();
            txtNewPass.Clear(); txtConfirm.Clear();
            cmbQuestion.SelectedIndex = -1;
        }
        catch (Exception ex)
        {
            lblMsg.ForeColor = Color.Red;
            lblMsg.Text = "DB error: " + ex.Message;
        }
    }
}

// ════════════════════════════════════════════════════════════
// 3. ABOUT
// ════════════════════════════════════════════════════════════
class AboutTab : TabPage
{
    public AboutTab()
    {
        Text = "About";
        Controls.Add(UI.MakeTitle("About the Program"));

        string[,] info =
        {
            { "System Name",  "School Information System" },
            { "Version",      "1.0.0" },
            { "Release Date", "2026" },
            { "Developer",    "IT Department" },
            { "Platform",     "Windows Forms (.NET)" },
            { "Database",     "MySQL / MariaDB (school)" },
            { "License",      "Proprietary" },
            { "Contact",      "support@school.local" }
        };

        int y = 60;
        for (int i = 0; i < info.GetLength(0); i++)
        {
            Controls.Add(new Label
            {
                Text = info[i, 0] + ":",
                Location = new Point(20, y),
                Width = 120,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });
            Controls.Add(new Label
            {
                Text = info[i, 1],
                Location = new Point(145, y),
                AutoSize = true
            });
            y += 28;
        }

        Controls.Add(new Label
        {

            Location = new Point(20, y + 10),
            AutoSize = true,
            ForeColor = Color.Gray
        });
    }
}

// ════════════════════════════════════════════════════════════
// 4. USER MANAGEMENT
// Covers: Add Account, Update Profile, Active/Inactive, List/Search
// ════════════════════════════════════════════════════════════
class UserManagementTab : TabPage
{
    DataGridView? grid;
    TextBox? txtSearch;
    ComboBox? cmbFilter;
    Label? lblStatus;

    public UserManagementTab()
    {
        Text = "User Management";
        Build();
        LoadAccounts();
    }

    void Build()
    {
        Controls.Add(UI.MakeLabel("Search:", 12, 16));
        txtSearch = UI.MakeTextBox(70, 13, 180);
        txtSearch.TextChanged += (s, e) => LoadAccounts();
        Controls.Add(txtSearch);

        Controls.Add(UI.MakeLabel("Status:", 268, 16));
        cmbFilter = UI.MakeCombo(318, 13, 100, new[] { "All", "Active", "Inactive" });
        cmbFilter.SelectedIndex = 0;
        cmbFilter.SelectedIndexChanged += (s, e) => LoadAccounts();
        Controls.Add(cmbFilter);

        var btnAdd = UI.MakeButton("+ Add", 435, 12, 70);
        btnAdd.Click += (s, e) => OpenAccountForm(null);
        Controls.Add(btnAdd);

        var btnEdit = UI.MakeButton("Edit", 513, 12, 60);
        btnEdit.Click += (s, e) => OpenEditForm();
        Controls.Add(btnEdit);

        var btnToggle = UI.MakeButton("Toggle Status", 581, 12, 110);
        btnToggle.Click += (s, e) => ToggleStatus();
        Controls.Add(btnToggle);

        grid = new DataGridView
        {
            Location = new Point(12, 45),
            Size = new Size(640, 330),
            ReadOnly = true,
            AllowUserToAddRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.Fixed3D
        };
        Controls.Add(grid);

        lblStatus = new Label { Location = new Point(12, 383), AutoSize = true, ForeColor = Color.Gray };
        Controls.Add(lblStatus);
    }

    public void LoadAccounts()
    {
        string search = txtSearch?.Text.Trim() ?? "";
        string filter = cmbFilter?.SelectedItem?.ToString() ?? "All";

        try
        {
            using var conn = Database.GetConnection();
            string sql = @"
                SELECT account_id, username, full_name, email, role,
                       IF(is_active=1,'Active','Inactive') AS status,
                       created_at
                FROM accounts
                WHERE (username LIKE @s OR full_name LIKE @s OR email LIKE @s)";

            if (filter == "Active") sql += " AND is_active = 1";
            if (filter == "Inactive") sql += " AND is_active = 0";
            sql += " ORDER BY created_at DESC";

            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@s", "%" + search + "%");

            using var adapter = new MySqlDataAdapter(cmd);
            var dt = new DataTable();
            adapter.Fill(dt);
            grid.DataSource = dt;

            void Rename(string col, string header)
            {
                if (grid.Columns.Contains(col))
                    grid.Columns[col]!.HeaderText = header;
            }
            Rename("account_id", "#");
            Rename("username", "Username");
            Rename("full_name", "Full Name");
            Rename("email", "Email");
            Rename("role", "Role");
            Rename("status", "Status");
            Rename("created_at", "Date Added");

            if (grid?.Columns.Contains("account_id") == true)
                grid.Columns["account_id"]!.FillWeight = 30;

            if (grid?.Columns.Contains("status") == true)
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    var cell = row.Cells["status"];
                    if (cell?.Value?.ToString() == "Active")
                        cell.Style.ForeColor = System.Drawing.Color.Green;
                    else if (cell != null)
                        cell.Style.ForeColor = Color.Red;
                }
            }

            lblStatus.Text = dt.Rows.Count + " record(s) found.";
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.Red;
            lblStatus.Text = "DB error: " + ex.Message;
        }
    }

    int? SelectedId()
    {
        if (grid == null) return null;
        if (grid.SelectedRows.Count == 0) return null;
        var val = grid.SelectedRows[0].Cells["account_id"].Value;
        return val == null ? null : Convert.ToInt32(val);
    }

    void OpenAccountForm(int? id)
    {
        using var dlg = new AccountForm(id);
        if (dlg.ShowDialog() == DialogResult.OK) LoadAccounts();
    }

    void OpenEditForm()
    {
        var id = SelectedId();
        if (id == null) { MessageBox.Show("Please select an account to edit.", "School IS"); return; }
        OpenAccountForm(id);
    }

    void ToggleStatus()
    {
        var id = SelectedId();
        if (id == null) { MessageBox.Show("Please select an account to toggle.", "School IS"); return; }

        if (grid == null) return;
        string cur = grid.SelectedRows[0].Cells["status"].Value?.ToString() ?? "";
        int newVal = cur == "Active" ? 0 : 1;
        string newLabel = newVal == 1 ? "Active" : "Inactive";

        if (MessageBox.Show("Set this account to " + newLabel + "?", "Confirm",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

        try
        {
            using var conn = Database.GetConnection();
            const string sql = "UPDATE accounts SET is_active=@a WHERE account_id=@id";
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@a", newVal);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.ExecuteNonQuery();
            if (lblStatus != null)
            {
                lblStatus.ForeColor = Color.Gray;
                lblStatus.Text = "Account set to " + newLabel + ".";
            }
            LoadAccounts();
        }
        catch (Exception ex)
        {
            lblStatus.ForeColor = Color.Red;
            lblStatus.Text = "DB error: " + ex.Message;
        }
    }
}

// ════════════════════════════════════════════════════════════
// ACCOUNT FORM  (Add / Edit dialog)
// ════════════════════════════════════════════════════════════
class AccountForm : Form
{
    readonly int? _id;
    TextBox txtUser, txtFull, txtEmail, txtPass, txtConf, txtAns;
    ComboBox cmbRole, cmbQ;
    Label lblMsg;

    public AccountForm(int? id)
    {
        _id = id;
        Text = id == null ? "Add Account" : "Edit Account";
        Size = new Size(460, 500);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false;

        Controls.Add(new Label
        {
            Text = id == null ? "Add New Account" : "Edit Account",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(12, 12)
        });

        int lx = 20, fx = 180, y = 55, gap = 36;

        void Row(string label, Control ctl)
        {
            Controls.Add(new Label { Text = label, Location = new Point(lx, y + 3), AutoSize = true });
            ctl.Location = new Point(fx, y);
            ctl.Width = 240;
            Controls.Add(ctl);
            y += gap;
        }

        txtUser = new TextBox();
        txtFull = new TextBox();
        txtEmail = new TextBox();
        cmbRole = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbRole.Items.AddRange(new[] { "Administrator", "Staff", "Viewer" });
        cmbQ = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        cmbQ.Items.AddRange(new object[]
        {
            "What is your mother's maiden name?",
            "What was the name of your first pet?",
            "What city were you born in?",
            "What is your elementary school name?"
        });
        txtAns = new TextBox();
        txtPass = new TextBox { PasswordChar = '*' };
        txtConf = new TextBox { PasswordChar = '*' };

        Row("Username:", txtUser);
        Row("Full Name:", txtFull);
        Row("Email:", txtEmail);
        Row("Role:", cmbRole);
        Row("Security Question:", cmbQ);
        Row("Security Answer:", txtAns);

        if (_id != null)
        {
            Controls.Add(new Label
            {
                Text = "(Leave blank to keep current password)",
                Location = new Point(fx, y),
                AutoSize = true,
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8)
            });
            y += 18;
        }

        Row("Password:", txtPass);
        Row(_id == null ? "Confirm Password:" : "Confirm New:", txtConf);

        var btnSave = UI.MakeButton(_id == null ? "Save" : "Update", fx, y + 6, 100);
        btnSave.Click += BtnSave_Click;
        Controls.Add(btnSave);

        var btnCancel = UI.MakeButton("Cancel", fx + 110, y + 6, 80);
        btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnCancel);

        lblMsg = new Label { Location = new Point(lx, y + 44), AutoSize = true, ForeColor = Color.Red };
        Controls.Add(lblMsg);

        if (id != null) LoadExisting(id.Value);
    }

    void LoadExisting(int id)
    {
        try
        {
            using var conn = Database.GetConnection();
            const string sql =
                "SELECT username, full_name, email, role, security_question, security_answer " +
                "FROM accounts WHERE account_id=@id";
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return;

            txtUser.Text = r["username"].ToString()!;
            txtFull.Text = r["full_name"].ToString()!;
            txtEmail.Text = r["email"].ToString()!;
            txtAns.Text = r["security_answer"].ToString()!;

            string role = r["role"].ToString()!;
            for (int i = 0; i < cmbRole.Items.Count; i++)
                if (cmbRole.Items[i]!.ToString() == role) { cmbRole.SelectedIndex = i; break; }

            string q = r["security_question"].ToString()!;
            for (int i = 0; i < cmbQ.Items.Count; i++)
                if (cmbQ.Items[i]!.ToString() == q) { cmbQ.SelectedIndex = i; break; }
        }
        catch (Exception ex) { lblMsg.Text = "Load error: " + ex.Message; }
    }

    void BtnSave_Click(object? sender, EventArgs e)
    {
        string user = txtUser.Text.Trim();
        string full = txtFull.Text.Trim();
        string email = txtEmail.Text.Trim();
        string pass = txtPass.Text.Trim();
        string conf = txtConf.Text.Trim();
        string ans = txtAns.Text.Trim();

        if (user == "" || full == "" || email == "" ||
            cmbRole.SelectedIndex < 0 || cmbQ.SelectedIndex < 0 || ans == "")
        { lblMsg.Text = "Please fill in all required fields."; return; }

        if (_id == null)
        {
            if (pass == "") { lblMsg.Text = "Password is required."; return; }
            if (pass != conf) { lblMsg.Text = "Passwords do not match."; return; }
            if (pass.Length < 6) { lblMsg.Text = "Password must be at least 6 characters."; return; }
        }
        else if (pass != "")
        {
            if (pass != conf) { lblMsg.Text = "Passwords do not match."; return; }
            if (pass.Length < 6) { lblMsg.Text = "Password must be at least 6 characters."; return; }
        }

        try
        {
            using var conn = Database.GetConnection();

            if (_id == null)
            {
                using var chk = new MySqlCommand(
                    "SELECT COUNT(*) FROM accounts WHERE username=@u", conn);
                chk.Parameters.AddWithValue("@u", user);
                if ((long)chk.ExecuteScalar()! > 0)
                { lblMsg.Text = "Username already exists."; return; }

                using var cmd = new MySqlCommand(
                    "INSERT INTO accounts " +
                    "(username, full_name, email, password_hash, role, " +
                    " security_question, security_answer, is_active, created_at) " +
                    "VALUES (@u,@f,@e,@p,@r,@q,@a,1,NOW())", conn);
                cmd.Parameters.AddWithValue("@u", user);
                cmd.Parameters.AddWithValue("@f", full);
                cmd.Parameters.AddWithValue("@e", email);
                cmd.Parameters.AddWithValue("@p", pass);
                cmd.Parameters.AddWithValue("@r", cmbRole.SelectedItem!.ToString());
                cmd.Parameters.AddWithValue("@q", cmbQ.SelectedItem!.ToString());
                cmd.Parameters.AddWithValue("@a", ans);
                cmd.ExecuteNonQuery();
            }
            else
            {
                using var chk = new MySqlCommand(
                    "SELECT COUNT(*) FROM accounts WHERE username=@u AND account_id<>@id", conn);
                chk.Parameters.AddWithValue("@u", user);
                chk.Parameters.AddWithValue("@id", _id.Value);
                if ((long)chk.ExecuteScalar()! > 0)
                { lblMsg.Text = "Username already taken."; return; }

                string upd =
                    "UPDATE accounts SET username=@u, full_name=@f, email=@e, " +
                    "role=@r, security_question=@q, security_answer=@a";
                if (pass != "") upd += ", password_hash=@p";
                upd += " WHERE account_id=@id";

                using var cmd = new MySqlCommand(upd, conn);
                cmd.Parameters.AddWithValue("@u", user);
                cmd.Parameters.AddWithValue("@f", full);
                cmd.Parameters.AddWithValue("@e", email);
                cmd.Parameters.AddWithValue("@r", cmbRole.SelectedItem!.ToString());
                cmd.Parameters.AddWithValue("@q", cmbQ.SelectedItem!.ToString());
                cmd.Parameters.AddWithValue("@a", ans);
                if (pass != "") cmd.Parameters.AddWithValue("@p", pass);
                cmd.Parameters.AddWithValue("@id", _id.Value);
                cmd.ExecuteNonQuery();
            }

            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex) { lblMsg.Text = "DB error: " + ex.Message; }
    }
}

// ════════════════════════════════════════════════════════════
// TRANSACTION 1: COURSE ENROLLMENT
// ════════════════════════════════════════════════════════════
class EnrollmentTab : TabPage
{
    ComboBox cmbStudent, cmbCourse, cmbNewDept;
    TextBox txtNewFirst, txtNewLast, txtNewEmail;
    DataGridView gridPreview;
    Label lblMsg, lblStudentMsg;

    public EnrollmentTab()
    {
        Text = "Enrollment";
        Controls.Add(UI.MakeTitle("Student Course Enrollment"));

        var gb = new GroupBox { Text = "New Enrollment Registration", Location = new Point(20, 50), Size = new Size(350, 180) };
        gb.Controls.Add(UI.MakeLabel("Select Student:", 15, 30));
        cmbStudent = UI.MakeCombo(120, 27, 210, new string[] { });
        gb.Controls.Add(cmbStudent);

        gb.Controls.Add(UI.MakeLabel("Select Course:", 15, 65));
        cmbCourse = UI.MakeCombo(120, 62, 210, new string[] { });
        gb.Controls.Add(cmbCourse);

        var btnEnroll = UI.MakeButton("Enroll Student", 120, 105, 150);
        btnEnroll.BackColor = Color.LightBlue;
        btnEnroll.Click += BtnEnroll_Click;
        gb.Controls.Add(btnEnroll);

        lblMsg = new Label { Location = new Point(120, 150), AutoSize = true };
        gb.Controls.Add(lblMsg);
        Controls.Add(gb);

        // Quick Add Student Section
        var gbStudent = new GroupBox { Text = "Quick Add Student", Location = new Point(20, 240), Size = new Size(350, 250) };
        gbStudent.Controls.Add(UI.MakeLabel("First Name:", 15, 30));
        txtNewFirst = UI.MakeTextBox(120, 27, 210);
        gbStudent.Controls.Add(txtNewFirst);

        gbStudent.Controls.Add(UI.MakeLabel("Last Name:", 15, 65));
        txtNewLast = UI.MakeTextBox(120, 62, 210);
        gbStudent.Controls.Add(txtNewLast);

        gbStudent.Controls.Add(UI.MakeLabel("Email:", 15, 100));
        txtNewEmail = UI.MakeTextBox(120, 97, 210);
        gbStudent.Controls.Add(txtNewEmail);

        gbStudent.Controls.Add(UI.MakeLabel("Dept:", 15, 135));
        cmbNewDept = UI.MakeCombo(120, 132, 210, new string[] { });
        gbStudent.Controls.Add(cmbNewDept);

        var btnAddStudent = UI.MakeButton("Create Student", 120, 170, 150);
        btnAddStudent.Click += BtnAddStudent_Click;
        gbStudent.Controls.Add(btnAddStudent);

        lblStudentMsg = new Label { Location = new Point(120, 210), AutoSize = true };
        gbStudent.Controls.Add(lblStudentMsg);
        Controls.Add(gbStudent);

        // Preview Grid
        Controls.Add(UI.MakeLabel("Recent Enrollments:", 390, 55));
        gridPreview = new DataGridView 
        { 
            Location = new Point(390, 75), 
            Size = new Size(420, 450), 
            ReadOnly = true, 
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.WhiteSmoke,
            RowHeadersVisible = false
        };
        Controls.Add(gridPreview);

        this.Enter += (s, e) => { RefreshLists(); LoadPreview(); };
    }

    void LoadPreview()
    {
        try {
            using var conn = Database.GetConnection();
            string sql = "SELECT CONCAT(s.lastname, ', ', s.firstname) as Student, GROUP_CONCAT(DISTINCT c.coursename SEPARATOR '; ') as 'Enrolled Courses' " +
                         "FROM enrollment e JOIN student s ON e.studentid = s.studentid " +
                         "JOIN course c ON e.courseid = c.courseid GROUP BY s.studentid ORDER BY s.lastname ASC";
            using var adapter = new MySqlDataAdapter(sql, conn);
            var dt = new DataTable();
            adapter.Fill(dt);
            gridPreview.DataSource = dt;
        } catch { }
    }

    void RefreshLists()
    {
        try {
            using var conn = Database.GetConnection();
            cmbStudent.Items.Clear();
            var cmdS = new MySqlCommand("SELECT studentid, CONCAT(lastname, ', ', firstname) as name FROM student ORDER BY lastname", conn);
            using (var r = cmdS.ExecuteReader()) while (r.Read()) cmbStudent.Items.Add(new { Id = r["studentid"], Name = r["name"].ToString() });
            cmbStudent.DisplayMember = "Name";
            
            cmbCourse.Items.Clear();
            var cmdC = new MySqlCommand("SELECT courseid, coursename FROM course", conn);
            using (var r = cmdC.ExecuteReader()) while (r.Read()) cmbCourse.Items.Add(new { Id = r["courseid"], Name = r["coursename"].ToString() });
            cmbCourse.DisplayMember = "Name";

            cmbNewDept.Items.Clear();
            var cmdD = new MySqlCommand("SELECT deptid, deptname FROM department ORDER BY deptname", conn);
            using (var r = cmdD.ExecuteReader()) while (r.Read()) cmbNewDept.Items.Add(new { Id = r["deptid"], Name = r["deptname"].ToString() });
            cmbNewDept.DisplayMember = "Name";
        } catch { }
    }

    void BtnAddStudent_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtNewFirst.Text) || string.IsNullOrWhiteSpace(txtNewLast.Text) || cmbNewDept.SelectedItem == null)
        {
            lblStudentMsg.ForeColor = Color.Red;
            lblStudentMsg.Text = "Fill in Name and Dept.";
            return;
        }

        try
        {
            using var conn = Database.GetConnection();
            const string sql = "INSERT INTO student (firstname, lastname, email, deptid) VALUES (@f, @l, @e, @d)";
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@f", txtNewFirst.Text.Trim());
            cmd.Parameters.AddWithValue("@l", txtNewLast.Text.Trim());
            cmd.Parameters.AddWithValue("@e", txtNewEmail.Text.Trim());
            cmd.Parameters.AddWithValue("@d", ((dynamic)cmbNewDept.SelectedItem).Id);
            cmd.ExecuteNonQuery();

            long newId = cmd.LastInsertedId;
            lblStudentMsg.ForeColor = Color.Green;
            lblStudentMsg.Text = "Student added!";
            
            // Refresh main list and auto-select the new student
            RefreshLists();
            foreach (var item in cmbStudent.Items)
            {
                if (((dynamic)item).Id.ToString() == newId.ToString())
                {
                    cmbStudent.SelectedItem = item;
                    break;
                }
            }

            txtNewFirst.Clear(); txtNewLast.Clear(); txtNewEmail.Clear();
        }
        catch (Exception ex)
        {
            lblStudentMsg.ForeColor = Color.Red;
            lblStudentMsg.Text = "Error: " + ex.Message;
        }
    }

    void BtnEnroll_Click(object? sender, EventArgs e)
    {
        if (cmbStudent.SelectedItem == null || cmbCourse.SelectedItem == null) { lblMsg.Text = "Select both student and course."; return; }
        try {
            using var conn = Database.GetConnection();

            // Prevention: Check if the student is already enrolled in this specific course
            const string checkSql = "SELECT COUNT(*) FROM enrollment WHERE studentid = @s AND courseid = @c";
            using var checkCmd = new MySqlCommand(checkSql, conn);
            checkCmd.Parameters.AddWithValue("@s", ((dynamic)cmbStudent.SelectedItem).Id);
            checkCmd.Parameters.AddWithValue("@c", ((dynamic)cmbCourse.SelectedItem).Id);
            if ((long)checkCmd.ExecuteScalar()! > 0)
            {
                lblMsg.ForeColor = Color.Red;
                lblMsg.Text = "Student is already enrolled in this course.";
                return;
            }

            var cmd = new MySqlCommand("INSERT INTO enrollment (studentid, courseid, grade) VALUES (@s, @c, 'F')", conn);
            cmd.Parameters.AddWithValue("@s", ((dynamic)cmbStudent.SelectedItem).Id);
            cmd.Parameters.AddWithValue("@c", ((dynamic)cmbCourse.SelectedItem).Id);
            cmd.ExecuteNonQuery();
            lblMsg.ForeColor = Color.Green;
            lblMsg.Text = "Enrollment successful!";
            LoadPreview();
        } catch (Exception ex) { lblMsg.Text = "Error: " + ex.Message; }
    }
}

// ════════════════════════════════════════════════════════════
// TRANSACTION 2: GRADE ENCODING
// ════════════════════════════════════════════════════════════
class GradeEncodingTab : TabPage
{
    ComboBox cmbStudent, cmbCourse, cmbGrade;
    DataGridView gridGrades;
    Label lblMsg;

    public GradeEncodingTab()
    {
        Text = "Grade Encoding";
        Controls.Add(UI.MakeTitle("Grade Encoding"));

        var gb = new GroupBox { Text = "Update Student Performance", Location = new Point(20, 50), Size = new Size(450, 200) };
        gb.Controls.Add(UI.MakeLabel("Student Name:", 15, 30));
        cmbStudent = UI.MakeCombo(120, 27, 310, new string[] { });
        cmbStudent.SelectedIndexChanged += (s, e) => LoadCoursesForStudent();
        gb.Controls.Add(cmbStudent);

        gb.Controls.Add(UI.MakeLabel("Select Course:", 15, 65));
        cmbCourse = UI.MakeCombo(120, 62, 310, new string[] { });
        gb.Controls.Add(cmbCourse);

        gb.Controls.Add(UI.MakeLabel("Final Grade:", 15, 100));
        cmbGrade = UI.MakeCombo(120, 97, 100, new[] { "A", "B", "C", "D", "F" });
        gb.Controls.Add(cmbGrade);

        var btnSave = UI.MakeButton("Save Grade", 120, 135, 120);
        btnSave.BackColor = Color.LightGreen;
        btnSave.Click += BtnSave_Click;
        gb.Controls.Add(btnSave);

        lblMsg = new Label { Location = new Point(250, 140), AutoSize = true };
        gb.Controls.Add(lblMsg);
        Controls.Add(gb);

        // Visual List of current grades
        Controls.Add(UI.MakeLabel("Class Record Overview:", 20, 260));
        gridGrades = new DataGridView
        {
            Location = new Point(20, 280),
            Size = new Size(790, 250),
            ReadOnly = true,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = Color.White,
            RowHeadersVisible = false
        };
        Controls.Add(gridGrades);

        this.Enter += (s, e) => { LoadStudents(); LoadGrades(); };
    }

    void LoadGrades()
    {
        try {
            using var conn = Database.GetConnection();
            string sql = "SELECT CONCAT(s.lastname, ', ', s.firstname) as 'Student', GROUP_CONCAT(DISTINCT CONCAT(c.coursename, ' (', e.grade, ')') SEPARATOR ' | ') as 'Grades' " +
                         "FROM enrollment e JOIN student s ON e.studentid = s.studentid " +
                         "JOIN course c ON e.courseid = c.courseid GROUP BY s.studentid ORDER BY s.lastname ASC";
            using var adapter = new MySqlDataAdapter(sql, conn);
            var dt = new DataTable();
            adapter.Fill(dt);
            gridGrades.DataSource = dt;
        } catch { }
    }

    void LoadStudents()
    {
        try {
            using var conn = Database.GetConnection();
            cmbStudent.Items.Clear();
            cmbCourse.Items.Clear();
            var sql = "SELECT studentid, CONCAT(lastname, ', ', firstname) as name FROM student ORDER BY lastname";
            var cmd = new MySqlCommand(sql, conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) cmbStudent.Items.Add(new { Id = r["studentid"], Name = r["name"].ToString() });
            cmbStudent.DisplayMember = "Name";
        } catch { }
    }

    void LoadCoursesForStudent()
    {
        cmbCourse.Items.Clear();
        if (cmbStudent.SelectedItem == null) return;
        try {
            using var conn = Database.GetConnection();
            var sql = "SELECT e.enrollmentid, c.coursename " +
                      "FROM enrollment e JOIN course c ON e.courseid = c.courseid " +
                      "WHERE e.studentid = @s ORDER BY c.coursename";
            var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@s", ((dynamic)cmbStudent.SelectedItem).Id);
            using var r = cmd.ExecuteReader();
            while (r.Read()) cmbCourse.Items.Add(new { Id = r["enrollmentid"], Name = r["coursename"].ToString() });
            cmbCourse.DisplayMember = "Name";
        } catch { }
    }

    void BtnSave_Click(object? sender, EventArgs e)
    {
        if (cmbCourse.SelectedItem == null || cmbGrade.SelectedIndex < 0) return;
        try {
            using var conn = Database.GetConnection();
            var cmd = new MySqlCommand("UPDATE enrollment SET grade = @g WHERE enrollmentid = @id", conn);
            cmd.Parameters.AddWithValue("@g", cmbGrade.SelectedItem?.ToString());
            cmd.Parameters.AddWithValue("@id", ((dynamic)cmbCourse.SelectedItem).Id);
            cmd.ExecuteNonQuery();
            lblMsg.Text = "Grade updated!";
            LoadGrades();
        } catch (Exception ex) { lblMsg.Text = "Error: " + ex.Message; }
    }
}

// ════════════════════════════════════════════════════════════
// TRANSACTION 3: INSTRUCTOR ASSIGNMENT
// ════════════════════════════════════════════════════════════
class InstructorAssignmentTab : TabPage
{
    ComboBox cmbCourse, cmbInstructor;
    Label lblMsg;
    public InstructorAssignmentTab()
    {
        Text = "Assignments";
        Controls.Add(UI.MakeTitle("Instructor Assignment"));
        Controls.Add(UI.MakeLabel("Course:", 20, 60));
        cmbCourse = UI.MakeCombo(150, 57, 250, new string[] { });
        Controls.Add(cmbCourse);
        Controls.Add(UI.MakeLabel("Instructor:", 20, 95));
        cmbInstructor = UI.MakeCombo(150, 92, 250, new string[] { });
        Controls.Add(cmbInstructor);
        var btnAssign = UI.MakeButton("Assign", 150, 130, 100);
        btnAssign.Click += (s, e) => {
            if (cmbCourse.SelectedItem == null || cmbInstructor.SelectedItem == null) return;
            try {
                using var conn = Database.GetConnection();
                var cmd = new MySqlCommand("UPDATE course SET instructorid=@i WHERE courseid=@c", conn);
                cmd.Parameters.AddWithValue("@i", ((dynamic)cmbInstructor.SelectedItem).Id);
                cmd.Parameters.AddWithValue("@c", ((dynamic)cmbCourse.SelectedItem).Id);
                cmd.ExecuteNonQuery();
                lblMsg.Text = "Instructor assigned successfully!";
            } catch (Exception ex) { lblMsg.Text = ex.Message; }
        };
        Controls.Add(btnAssign);
        lblMsg = new Label { Location = new Point(150, 170), AutoSize = true };
        Controls.Add(lblMsg);
        this.Enter += (s, e) => {
            try {
                using var conn = Database.GetConnection();
                cmbCourse.Items.Clear();
                var cmdC = new MySqlCommand("SELECT courseid, coursename FROM course", conn);
                using (var r = cmdC.ExecuteReader()) while (r.Read()) cmbCourse.Items.Add(new { Id = r["courseid"], Name = r["coursename"].ToString() });
                cmbCourse.DisplayMember = "Name";
                cmbInstructor.Items.Clear();
                var cmdI = new MySqlCommand("SELECT instructorid, CONCAT(firstname, ' ', lastname) as name FROM instructor", conn);
                using (var r = cmdI.ExecuteReader()) while (r.Read()) cmbInstructor.Items.Add(new { Id = r["instructorid"], Name = r["name"].ToString() });
                cmbInstructor.DisplayMember = "Name";
            } catch { }
        };
    }
}

// ════════════════════════════════════════════════════════════
// REPORT GENERATION MODULE
// ════════════════════════════════════════════════════════════
class ReportModuleTab : TabPage
{
    DataGridView grid;
    ComboBox cmbReportType;
    public ReportModuleTab()
    {
        Text = "Reports";
        Controls.Add(UI.MakeTitle("System Reports"));
        Controls.Add(UI.MakeLabel("Select Report:", 20, 60));
        cmbReportType = UI.MakeCombo(150, 57, 250, new[] { "Enrollment Summary", "Grade Distribution", "Instructor Workload" });
        cmbReportType.SelectedIndex = 0;
        cmbReportType.SelectedIndexChanged += (s, e) => LoadData();
        Controls.Add(cmbReportType);
        grid = new DataGridView { Location = new Point(20, 100), Size = new Size(780, 350), ReadOnly = true, BackgroundColor = Color.White, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
        Controls.Add(grid);
        var btnExport = UI.MakeButton("Export to Excel", 650, 57, 150);
        btnExport.Click += (s, e) => ExportReport();
        Controls.Add(btnExport);
        this.Enter += (s, e) => LoadData();
    }
    void LoadData()
    {
        string sql = cmbReportType.SelectedIndex switch {
            0 => "SELECT studentname as 'Student', courses as 'Enrolled Courses' FROM vwstudentenrollments",
            1 => "SELECT grade as 'Grade', COUNT(*) as 'Total' FROM enrollment WHERE grade IS NOT NULL GROUP BY grade",
            _ => "SELECT instructorname as 'Instructor', totalcourses as 'Course Count' FROM vwinstructorload"
        };
        try {
            using var conn = Database.GetConnection();
            using var adapter = new MySqlDataAdapter(sql, conn);
            var dt = new DataTable();
            adapter.Fill(dt);
            grid.DataSource = dt;
        } catch { }
    }
    void ExportReport()
    {
        var dt = grid.DataSource as DataTable;
        if (dt == null || dt.Rows.Count == 0) return;

        try
        {
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            string filePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            );

            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                // Data Report Sheet
                var ws1 = package.Workbook.Worksheets.Add("Data Report");
                ws1.Cells[1, 1].Value = "Burg & Roy National Highschool";
                ws1.Cells[1, 1].Style.Font.Bold = true;
                ws1.Cells[1, 1].Style.Font.Size = 16;
                ws1.Cells[2, 1].Value = "[LOGO PLACEHOLDER]";
                ws1.Cells[4, 1].Value = cmbReportType.SelectedItem?.ToString();

                // Add column headers
                for (int i = 0; i < dt.Columns.Count; i++)
                    ws1.Cells[6, i + 1].Value = dt.Columns[i].ColumnName;

                // Add data rows
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int j = 0; j < dt.Columns.Count; j++)
                        ws1.Cells[i + 7, j + 1].Value = dt.Rows[i][j]?.ToString() ?? "";

                int lastRow = dt.Rows.Count + 6;
                ws1.Cells[lastRow + 2, 1].Value = "Prepared by: " + (Session.CurrentUser ?? "System");
                ws1.Cells[lastRow + 3, 1].Value = "________________________";
                ws1.Column(1).Width = 20;

                // Chart Sheet - Create a meaningful chart from report data
                var wsChart = package.Workbook.Worksheets.Add("Chart");
                
                if (dt.Rows.Count > 0)
                {
                    var labels = new List<string>();
                    var values = new List<double>();
                    string xHeader = "Category";
                    string yHeader = "Value";

                    if (cmbReportType.SelectedIndex == 0)
                    {
                        // Updated Enrollment Summary: Show number of courses per student
                        xHeader = "Student";
                        yHeader = "Number of Courses";
                    }
                    
                    if (labels.Count == 0)
                    {
                        xHeader = dt.Columns[0].ColumnName;
                        yHeader = dt.Columns.Count > 1 ? dt.Columns[1].ColumnName : "Value";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            labels.Add(dt.Rows[i][0]?.ToString() ?? "");
                            
                            var valStr = dt.Rows[i][1]?.ToString() ?? "";
                            if (double.TryParse(valStr, out double numericValue))
                                values.Add(numericValue);
                            else if (cmbReportType.SelectedIndex == 0)
                                // For the courses list, count the semicolons to determine count
                                values.Add(valStr.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).Length);
                            else
                                values.Add(0);
                        }
                    }

                    // Write chart table
                    wsChart.Cells[1, 1].Value = xHeader;
                    wsChart.Cells[1, 2].Value = yHeader;
                    for (int i = 0; i < labels.Count; i++)
                    {
                        wsChart.Cells[i + 2, 1].Value = labels[i];
                        wsChart.Cells[i + 2, 2].Value = values[i];
                    }

                    wsChart.Column(1).Width = 30;
                    wsChart.Column(2).Width = 15;

                    // Create a column chart with category labels
                    var chart = wsChart.Drawings.AddChart("DataChart", OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered);
                    chart.Title.Text = cmbReportType.SelectedItem?.ToString() ?? "Report Chart";
                    var xRange = wsChart.Cells[2, 1, labels.Count + 1, 1];
                    var yRange = wsChart.Cells[2, 2, labels.Count + 1, 2];
                    chart.Series.Add(yRange, xRange);

                    chart.SetPosition(1, 0, 3, 0);
                    chart.SetSize(800, 400);
                    chart.Legend.Position = OfficeOpenXml.Drawing.Chart.eLegendPosition.Right;
                }

                // Summary Sheet
                var ws2 = package.Workbook.Worksheets.Add("Summary");
                ws2.Cells[1, 1].Value = "Total Records:";
                ws2.Cells[1, 2].Value = dt.Rows.Count;
                ws2.Cells[2, 1].Value = "Columns:";
                ws2.Cells[2, 2].Value = dt.Columns.Count;
                ws2.Cells[3, 1].Value = "Report Type:";
                ws2.Cells[3, 2].Value = cmbReportType.SelectedItem?.ToString() ?? "Unknown";
                ws2.Cells[4, 1].Value = "Generated:";
                ws2.Cells[4, 2].Value = DateTime.Now;
                ws2.Column(1).Width = 20;

                package.SaveAs(new FileInfo(filePath));
                MessageBox.Show($"Report generated:\n{filePath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Export failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
