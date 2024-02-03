using System;
using System.Data;
using System.Windows.Forms;
using Npgsql;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;

namespace OnlineStore
{
    public partial class Form1 : Form
    {

        private string id = "";
        private int intRow = 0;
        private string nameTable = "worker";       

        public Form1()
        {
            InitializeComponent();                       
            loadData();
            resetMe();
        }

        private void resetMe()
        {            
            this.id = string.Empty;

            Name_WBox.Text = "";
            Surname_WBox.Text = "";
            Otchestvo_WBox.Text = "";
            Login_WBox.Text = "";
            Parol_WBox.Text = "";

            Diagnoz_RBox.Text = "";

            Name_PatBox.Text = "";
            Surname_PatBox.Text = "";
            Otchestvo_PatBox.Text = "";
            LgotaBox.Text = "";

            moneyProductNud.Value = 0;
            

            workerUpdateButton.Text = "Update ()";
            workerDeleteButton.Text = "Delete ()";
            recipeUpdateButton.Text = "Update ()";
            recipeDeleteButton.Text = "Delete ()";
            patientDeleteButton.Text = "Delete ()";
            patientUpdateButton.Text = "Update ()";
            chetDeleteButton.Text = "Delete ()";
            chetUpdateButton.Text = "Update ()";
            zakazDeleteButton.Text = "Delete ()";
            zakazUpdateButton.Text = "Update ()";
            selectDoctorLb_Update();
            selectPreparatLb_Update();
            selectRecipeLb_Update();
            selectWorkeLb_Update();
            selectPatientLb_Update();
            selectChetLb_Update();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadData();
        }

        private void loadData()
        {
            CRUD.sql = $"SELECT * FROM {nameTable}";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);

            if (dt.Rows.Count > 0)
            {
                intRow = Convert.ToInt32(dt.Rows.Count.ToString());
            }
            else
            {
                intRow = 0;
            }

            toolStripStatusLabel1.Text = "Number of row(s): " + intRow.ToString();

            DataGridView dgv1 = dataGridView1;

            dgv1.MultiSelect = false;
            dgv1.AutoGenerateColumns = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dgv1.DataSource = dt;

            dgv1.Columns[0].Width = 55;
        }
        // добавление работников
        private void customerInsertButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Name_WBox.Text.Trim()) 
                || string.IsNullOrEmpty(Surname_WBox.Text.Trim())
                || string.IsNullOrEmpty(Otchestvo_WBox.Text.Trim())
                || string.IsNullOrEmpty(Login_WBox.Text.Trim())
                || string.IsNullOrEmpty(Parol_WBox.Text.Trim()))
            {
                MessageBox.Show("Заполните все поля.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DateTime startDate = BDate_W.Value;

            CRUD.sql = $"INSERT INTO worker( surname, name, otchestvo, date_of_bir, login, parol) VALUES( @firstName,@lastName, @middleName, @1, @login, @parol)";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("lastName", Surname_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("firstName", Name_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("middleName", Otchestvo_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("1", startDate);
            CRUD.cmd.Parameters.AddWithValue("login", Login_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("parol", Parol_WBox.Text.Trim());
            if (CRUD.PerformCRUD(CRUD.cmd) != null)           
                MessageBox.Show("Данные сохранены.", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // обновление работников
        private void customerUpdateButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Пожалуйста, выберите элемент", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (string.IsNullOrEmpty(Surname_WBox.Text.Trim())
                || string.IsNullOrEmpty(Name_WBox.Text.Trim()) ||
                string.IsNullOrEmpty(Otchestvo_WBox.Text.Trim())
                 || string.IsNullOrEmpty(Login_WBox.Text.Trim())
                 || string.IsNullOrEmpty(Parol_WBox.Text.Trim()))
            {
                MessageBox.Show("Пожалуйста, заполните все поля", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            DateTime startDate = BDate_W.Value;
            CRUD.sql = $"UPDATE worker SET surname = @lastName, name = @firstName, otchestvo = @middleName, date_of_bir = @1, login = @log, parol = @parol  WHERE  id_worker  = @id::integer";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("lastName", Surname_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("firstName", Name_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("middleName", Otchestvo_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("1", startDate);
            CRUD.cmd.Parameters.AddWithValue("log", Login_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("parol", Parol_WBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("id", this.id);                                    

            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Обновлено", "Успешно",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // удаление работников
        private void deleteButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Выберите элементы, поле не может быть пустым!", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("Вы действительно хотите удалить?", "Удаление",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            {
                if (nameTable == "worker")
                {
                    CRUD.sql = $"DELETE FROM {nameTable} WHERE id_worker = @id::integer";
                }
                else  if (nameTable == "recipe")
                    {
                        CRUD.sql = $"DELETE FROM {nameTable} WHERE id_recipe = @id::integer";
                    }
                else if (nameTable == "patient")
                {
                    CRUD.sql = $"DELETE FROM {nameTable} WHERE id_patient = @id::integer";
                }
                else if (nameTable == "chet")
                {
                    CRUD.sql = $"DELETE FROM {nameTable} WHERE id_chet = @id::integer";
                }
                else if (nameTable == "zakaz")
                {
                    CRUD.sql = $"DELETE FROM {nameTable} WHERE id_zakaz = @id::integer";
                }

                CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
                CRUD.cmd.Parameters.Clear();
                CRUD.cmd.Parameters.AddWithValue("id", this.id);
                if (CRUD.PerformCRUD(CRUD.cmd) != null)
                    MessageBox.Show("Удалено.", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                loadData();
                resetMe();
            }

        }
        
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex != -1)
            {
                DataGridView dgv1 = dataGridView1;

                this.id = Convert.ToString(dgv1.CurrentRow.Cells[0].Value);
                workerUpdateButton.Text = "Update (" + this.id + ")";
                workerDeleteButton.Text = "Delete (" + this.id + ")";
                recipeUpdateButton.Text = "Update (" + this.id + ")";
                recipeDeleteButton.Text = "Delete (" + this.id + ")";
                patientDeleteButton.Text = "Delete (" + this.id + ")";
                patientUpdateButton.Text = "Update (" + this.id + ")";
                chetDeleteButton.Text = "Delete (" + this.id + ")";
                chetUpdateButton.Text = "Update (" + this.id + ")";
                zakazDeleteButton.Text = "Delete (" + this.id + ")";
                zakazUpdateButton.Text = "Update (" + this.id + ")";
            }

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            nameTable = tabControl1.TabPages[tabControl1.SelectedIndex].Text;
            loadData();
            resetMe();
        }
        // Вывод имен,названий для наглядности во всех вкладках
        private void selectDoctorLb_Update()
        {
            CRUD.sql = $"SELECT * FROM doctor";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            selectDoctorLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectDoctorLb.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}");    
            Id_DocBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_DocBox.Items.Add($"{dt.Rows[i][0]}");
        }
        private void selectPreparatLb_Update()
        {
            CRUD.sql = $"SELECT * FROM preparat";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            selectPreparatLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectPreparatLb.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}");
            selectPreparatLb1.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectPreparatLb1.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}");
            Id_PrepBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_PrepBox.Items.Add($"{dt.Rows[i][0]}");
            Id_PreparateBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_PreparateBox.Items.Add($"{dt.Rows[i][0]}");

        }
        private void selectRecipeLb_Update()
        {
            CRUD.sql = $"SELECT * FROM recipe";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            selectRecipeLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectRecipeLb.Items.Add($"{dt.Rows[i][0]}");
            Id_RecipeBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_RecipeBox.Items.Add($"{dt.Rows[i][0]}");
        }
        private void selectChetLb_Update()
        {
            CRUD.sql = $"SELECT * FROM chet";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            Id_ChetBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_ChetBox.Items.Add($"{dt.Rows[i][0]}");
        }
        private void selectWorkeLb_Update()
        {
            CRUD.sql = $"SELECT * FROM worker";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            selectWorkerLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectWorkerLb.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}" +
                    $" {dt.Rows[i][2]} {dt.Rows[i][3]}");
            Id_WBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_WBox.Items.Add($"{dt.Rows[i][0]}");
        }
        private void selectPatientLb_Update()
        {
            CRUD.sql = $"SELECT * FROM patient";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            System.Data.DataTable dt = CRUD.PerformCRUD(CRUD.cmd);
            selectPatientLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectPatientLb.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}" +
                    $" {dt.Rows[i][2]} {dt.Rows[i][3]}");
            Id_PatBox.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                Id_PatBox.Items.Add($"{dt.Rows[i][0]}");
            selectPatLb.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
                selectPatLb.Items.Add($"{dt.Rows[i][0]} : {dt.Rows[i][1]}" +
                    $" {dt.Rows[i][2]} {dt.Rows[i][3]}");
        }
      


        // вторая кнопка отправки отчета
        private void reportOrdersBtn_Click(object sender, EventArgs e)
        {
            if (selectPatLb.SelectedItems.Count <= 0)
                return;

            string customers = "";

            for (int i = 0; i < selectPatLb.SelectedItems.Count; i++)
                customers += selectPatLb.SelectedItems[i].ToString()[0] + ", ";            

            string ids = customers.Remove(customers.Length - 2);
            string startDate = ordersFromDtp.Value.Date.ToString("yyyy-MM-dd");
            string endDate = ordersToDtp.Value.Date.ToString("yyyy-MM-dd");            

            CRUD.sql = $"SELECT * FROM get_customer_orders(ARRAY[{ids}]," +
                $" '{startDate}'::DATE,'{endDate}'::DATE)";

            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();

            SaveDataTableToExcelAndFillDatagrid(CRUD.PerformCRUD(CRUD.cmd));            
        }

        private void SaveDataTableToExcelAndFillDatagrid(System.Data.DataTable table)
        {
            DataGridView dgv1 = dataGridView1;

            dgv1.MultiSelect = false;
            dgv1.AutoGenerateColumns = true;
            dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dgv1.DataSource = table;

            dgv1.Columns[0].Width = 55;

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            var worksheet = workbook.ActiveSheet;

            // Записываем заголовки столбцов
            for (int i = 0; i < table.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1] = table.Columns[i].ColumnName;
            }

            // Записываем данные
            for (int row = 0; row < table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1] = table.Rows[row][col];
                }
            }

            // Сохраняем файл
            var saveFileDialoge = new SaveFileDialog();
            saveFileDialoge.FileName = "otchet";
            saveFileDialoge.DefaultExt = ".xlsx";
            if (saveFileDialoge.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialoge.FileName);
            }
            else return;
            workbook.Close();
            excelApp.Quit();
        }

        // добавление recipe
        private void recipeInsertButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Id_DocBox.Text) || string.IsNullOrEmpty(Id_PrepBox.Text))
            {
                MessageBox.Show("Выберите id", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (string.IsNullOrEmpty(Diagnoz.Text.Trim()))
            {
                MessageBox.Show("Введите диагноз", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            CRUD.sql = $"INSERT INTO recipe (id_doctor, diagnos, id_prep ) VALUES(@1, @2, @3)";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", int.Parse(Id_DocBox.Text));
            CRUD.cmd.Parameters.AddWithValue("2", Diagnoz_RBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_PrepBox.Text));
            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные сохранены", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }

        private void recipeUpdateButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Выберите id для изменения", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (string.IsNullOrEmpty(Diagnoz.Text.Trim()))
            {
                MessageBox.Show("Введите наименование диагноза", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (Id_PrepBox.Items.Count <= 0)
                return;
            if (Id_DocBox.Items.Count <= 0)
                return;
            string customers = "";

                customers += Id_DocBox.Items.ToString()[0];

            

            CRUD.sql = $"UPDATE recipe SET id_doctor = @1, diagnos = @2, id_prep = @3 WHERE id_recipe = @id::integer";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", int.Parse(Id_DocBox.Text));
            CRUD.cmd.Parameters.AddWithValue("2", Diagnoz_RBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_PrepBox.Text));
            CRUD.cmd.Parameters.AddWithValue("id", this.id);            

            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные обновлены", "Успешно",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // добавление patient
        private void patientInsertButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Name_PatBox.Text.Trim())
                || string.IsNullOrEmpty(Surname_PatBox.Text.Trim())
                || string.IsNullOrEmpty(Otchestvo_PatBox.Text.Trim())
                || string.IsNullOrEmpty(LgotaBox.Text.Trim()))
            {
                MessageBox.Show("Заполните все поля.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (string.IsNullOrEmpty(Id_RecipeBox.Text))
            {
                MessageBox.Show("Выберите id", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DateTime startDate = BDate_P.Value;

            CRUD.sql = $"INSERT INTO patient( surname, name, otchestvo, date_of_bir, id_recipe, lgot) VALUES( @lastName, @firstName, @middleName, @1, @2, @lgot)";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("lastName", Surname_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("firstName", Name_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("middleName", Otchestvo_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("1", startDate);
            CRUD.cmd.Parameters.AddWithValue("2", int.Parse(Id_RecipeBox.Text));
            CRUD.cmd.Parameters.AddWithValue("lgot", LgotaBox.Text.Trim());
            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные сохранены.", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // обновление работников
        private void patientUpdateButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Пожалуйста, выберите элемент", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (string.IsNullOrEmpty(Name_PatBox.Text.Trim())
               || string.IsNullOrEmpty(Surname_PatBox.Text.Trim())
               || string.IsNullOrEmpty(Otchestvo_PatBox.Text.Trim())
               || string.IsNullOrEmpty(LgotaBox.Text.Trim()))
            {
                MessageBox.Show("Пожалуйста, заполните все поля", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            DateTime startDate = BDate_W.Value;
            CRUD.sql = $"UPDATE patient SET surname = @lastName, name = @firstName, otchestvo = @middleName, date_of_bir = @1, id_recipe = @2, lgot = @lgot  WHERE  id_patient  = @id::integer";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("lastName", Surname_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("firstName", Name_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("middleName", Otchestvo_PatBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("1", startDate);
            CRUD.cmd.Parameters.AddWithValue("2", int.Parse(Id_RecipeBox.Text));
            CRUD.cmd.Parameters.AddWithValue("lgot", LgotaBox.Text.Trim());
            CRUD.cmd.Parameters.AddWithValue("id", this.id);

            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Обновлено", "Успешно",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // добавление zakaz
        private void zakazInsertButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Id_PreparateBox.Text) || string.IsNullOrEmpty(Id_PatBox.Text) || string.IsNullOrEmpty(Id_ChetBox.Text))
            {
                MessageBox.Show("Выберите id", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            CRUD.sql = $"INSERT INTO zakaz (id_prep, id_patient, id_chet ) VALUES(@1, @2, @3)";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", int.Parse(Id_PreparateBox.Text));
            CRUD.cmd.Parameters.AddWithValue("2", int.Parse(Id_PatBox.Text));
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_ChetBox.Text));
            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные сохранены", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }

        private void zakazUpdateButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Выберите id для изменения", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (Id_PatBox.Items.Count <= 0)
                return;
            if (Id_PrepBox.Items.Count <= 0)
                return;
            if (Id_PreparateBox.Items.Count <= 0)
                return;


            CRUD.sql = $"UPDATE zakaz SET id_prep = @1, id_patient = @2, id_chet = @3 WHERE id_zakaz = @id::integer";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", int.Parse(Id_PreparateBox.Text));
            CRUD.cmd.Parameters.AddWithValue("2", int.Parse(Id_PatBox.Text));
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_ChetBox.Text));
            CRUD.cmd.Parameters.AddWithValue("id", this.id);

            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные обновлены", "Успешно",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }
        // добавление chet
        private void chetInsertButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Id_WBox.Text))
            {
                MessageBox.Show("Выберите id", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            DateTime startDate = CData.Value;
            CRUD.sql = $"INSERT INTO chet (price, date, id_worker ) VALUES(@1, @2, @3)";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", moneyProductNud.Value);
            CRUD.cmd.Parameters.AddWithValue("2", startDate);
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_WBox.Text));
            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные сохранены", "Успешно",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }

        private void chetUpdateButton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.id))
            {
                MessageBox.Show("Выберите id для изменения", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (Id_WBox.Items.Count <= 0)
                return;

            DateTime startDate = CData.Value;
            CRUD.sql = $"UPDATE chet SET price = @1, date = @2, id_worker = @3 WHERE id_chet = @id::integer";
            CRUD.cmd = new NpgsqlCommand(CRUD.sql, CRUD.con);
            CRUD.cmd.Parameters.Clear();
            CRUD.cmd.Parameters.AddWithValue("1", moneyProductNud.Value);
            CRUD.cmd.Parameters.AddWithValue("2", startDate);
            CRUD.cmd.Parameters.AddWithValue("3", int.Parse(Id_WBox.Text));
            CRUD.cmd.Parameters.AddWithValue("id", this.id);

            if (CRUD.PerformCRUD(CRUD.cmd) != null)
                MessageBox.Show("Данные обновлены", "Успешно",
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            loadData();
            resetMe();
        }

    }
}