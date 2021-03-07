using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//使用Excel需要用到的命名空间
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace ProgrammingDesign
{
    public partial class MainForm : Form
    {
        public int times = 0;//添加一个变量用于记录DataGridView控件的行列数并赋初值为0
        public int search = -1;//添加一个变量用于记录查询时所使用的方法并赋初值为-1(用以表示待机状态)
        public MainForm()
        {
            InitializeComponent();

        }

        private void 退出EscToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult result;//创建变量用于接收对话框中的返回值
            result = MessageBox.Show("确定退出程序吗?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //新建一个提示对话框包括按钮是否，并将返回值赋给result
            if (result == DialogResult.Yes)//判断对话框返回值，若为是则关闭程序
            {
                Application.Exit();//关闭程序
            }
        }

        private void 录入ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //判断是否处于查询模式，若是，则禁止进行信息录入
            if (search != -1)
            {
                MessageBox.Show("查询模式下不能进行信息录入！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //判断进行录入时，学号文本框和姓名文本框中是否有信息，若为空，则禁止录入(学号唯一，必须输入;学号+姓名可以用于确定身份，所以必须录入)
            if (SNumTextbox.Text == "" && NameTextbox.Text == "")
            {
                MessageBox.Show("学号/姓名为空，禁止录入无法确定身份的信息！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (SNumTextbox.Text.Length < 10)
            {
                MessageBox.Show("学号输入不合法，学号应为10位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (CNumTextbox.Text != "")
            {
                if (CNumTextbox.Text.Length < 3)
                {
                    MessageBox.Show("班号输入不合法，应为3-4位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            //未处于查询模式时可以进行信息录入
            times = StudentData.Rows.Add();//创建新行并返回新行的行号赋值给times
            for (int i = 0; i < times; i++)//遍历搜索
            {
                if (SNumTextbox.Text == StudentData.Rows[i].Cells[0].Value.ToString())
                {
                    MessageBox.Show("非法的重复录入！");
                    StudentData.Rows.RemoveAt(times);//重复录入则删除创建的空行
                    times--;
                    return;
                }
            }
            //一开始的时候，times=StudentData.Rows.Add();的语句是放在这个位置的，在测试中发现放在这个位置时第一二行无法判断是否重复
            //(因为初始的times都是0所以没有进入for循环)于是将创建新行的语句放在录入事件的最前端，然后在if判断句的循环体中将语句开始创建的新行删除
            //从而达到录入时对学号是否重复的判断(毕竟学号是不会重复的，姓名之类的都可以重复)
            StudentData.Rows[times].Cells[0].Value = SNumTextbox.Text;//向dataGridView控件的第times行学号列的单元格添加元素
            StudentData.Rows[times].Cells[1].Value = NameTextbox.Text;//向dataGridView控件的第times行姓名列的单元格添加元素
            StudentData.Rows[times].Cells[2].Value = GenderTextbox.Text;//向dataGridView控件的第times行性别列的单元格添加元素
            StudentData.Rows[times].Cells[3].Value = AgeTextbox.Text;//向dataGridView控件的第times行年龄列的单元格添加元素
            StudentData.Rows[times].Cells[4].Value = CNumTextbox.Text;//向dataGridView控件的第times行班号列的单元格添加元素
            StudentData.Rows[times].Cells[5].Value = SpecialityTextbox.Text;//向dataGridView控件的第times行专业列的单元格添加元素
            StudentData.Rows[times].Cells[6].Value = DepartmentTextbox.Text;//向dataGridView控件的第times行系别列的单元格添加元素

            //信息录入后将所有文本框的文本清空，减少手动操作
            SNumTextbox.Text = "";
            NameTextbox.Text = "";
            GenderTextbox.Text = "";
            AgeTextbox.Text = "";
            CNumTextbox.Text = "";
            SpecialityTextbox.Text = "";
            DepartmentTextbox.Text = "";
        }

        private void SNumTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许学号栏输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//仅允许学号栏输入数字0-9
                {
                    e.Handled = true;
                }
            }
        }

        private void AgeTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许年龄栏输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//仅允许年龄栏输入数字0-9
                {
                    e.Handled = true;
                }
            }
        }

        private void NameTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许姓名栏输入退格键
            {
                if (!char.IsLetter(e.KeyChar))//禁止姓名栏输入数字及符号
                {
                    e.Handled = true;
                }
            }
        }

        private void GenderTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许性别栏输入退格键
            {
                if (!char.IsLetter(e.KeyChar))//禁止性别栏输入数字及符号
                {
                    e.Handled = true;
                }
                else
                {
                    if ((e.KeyChar >= 'a') && (e.KeyChar <= 'z'))//禁止性别栏输入小写字母a-z
                    {
                        e.Handled = true;
                    }
                    else if ((e.KeyChar >= 'A') && (e.KeyChar <= 'Z'))//禁止性别栏输入大写字母A-Z
                    {
                        e.Handled = true;
                    }
                }
            }
        }

        private void CNumTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许班号栏输入退格键
            {
                if ((e.KeyChar < '0') || (e.KeyChar > '9'))//仅允许班号栏输入数字0-9
                {
                    e.Handled = true;
                }
            }
        }

        private void SpecialityTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许专业栏输入退格键
            {
                if (!char.IsLetter(e.KeyChar))//禁止专业栏输入数字及符号
                {
                    e.Handled = true;
                }
                else
                {
                    if ((e.KeyChar >= 'a') && (e.KeyChar <= 'z'))//禁止专业栏输入小写字母a-z
                    {
                        e.Handled = true;
                    }
                    else if ((e.KeyChar >= 'A') && (e.KeyChar <= 'Z'))//禁止专业栏输入大写字母A-Z
                    {
                        e.Handled = true;
                    }
                }
            }
        }

        private void DepartmentTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b')//允许系别栏输入退格键
            {
                if (!char.IsLetter(e.KeyChar))//禁止系别栏输入数字及符号
                {
                    e.Handled = true;
                }
                else
                {
                    if ((e.KeyChar >= 'a') && (e.KeyChar <= 'z'))//禁止系别栏输入小写字母a-z
                    {
                        e.Handled = true;
                    }
                    else if ((e.KeyChar >= 'A') && (e.KeyChar <= 'Z'))//禁止系别栏输入大写字母A-Z
                    {
                        e.Handled = true;
                    }
                }
            }
        }
        private void 按学号ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可查询");
                return;
            }
            else
            {
                MessageBox.Show("您已进入查询模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                search = 0;//search等于0表示按学号查询模式
                if (SNumTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于查询的学号");
                    MessageBox.Show("查询模式已退出");
                    search = -1;//查询模式退出时将search的值还原为-1
                    return;
                }
                else
                {
                    //这一层for循环才是整个的查询操作，前面的if条件句均为是否可进行查询的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的学号在数据列表中搜索信息，相关信息设置为可见，无关信息设置为不可见
                        if (SNumTextbox.Text == StudentData.Rows[i].Cells[search].Value.ToString())
                        {
                            StudentData.Rows[i].Visible = true;//搜索到的相关项设置为可见
                        }
                        else
                        {
                            StudentData.Rows[i].Visible = false;//无关项设置为不可见
                        }
                    }

                    //查询操作结束后再用一层for循环遍历DataGridView的所有行用于判断查询结束后是否有可见行
                    bool temp = true;//创建一个临时的布尔值用储存判断状态
                    for (int i = 0; i <= times; i++)
                    {
                        //对DataGridView控件的所有行进行判断，若存在可见项，将temp值置为false
                        if (StudentData.Rows[i].Visible != false)
                        {
                            temp = false;
                        }

                    }
                    if (temp != false)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        for (int j = 0; j <= times; j++)
                        {
                            StudentData.Rows[j].Visible = true;//将DataGridView的所有行的可见性还原
                        }
                        MessageBox.Show("查询模式已退出");
                        search = -1;
                        return;//这里没有清除学号文本框的文本是为了让用户检查是否有输入错误(其实就是懒得写)
                    }
                    SNumTextbox.Text = "";//同样为了减少操作，将学号文本框的文本清除
                }
            }

        }


        private void 按姓名ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可查询");
                return;
            }
            else
            {
                MessageBox.Show("您已进入查询模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                search = 1;//search等于1表示按姓名查询模式
                if (NameTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于查询的姓名");
                    MessageBox.Show("查询模式已退出");
                    search = -1;//查询模式退出时将search的值还原为-1
                    return;
                }
                else
                {
                    //这一层for循环才是整个的查询操作，前面的if条件句均为是否可进行查询的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的姓名在数据列表中搜索信息，相关信息设置为可见，无关信息设置为不可见
                        if (NameTextbox.Text == StudentData.Rows[i].Cells[search].Value.ToString())
                        {
                            StudentData.Rows[i].Visible = true;//搜索到的相关项设置为可见
                        }
                        else
                        {
                            StudentData.Rows[i].Visible = false;//无关项设置为不可见
                        }
                    }

                    //查询操作结束后再用一层for循环遍历DataGridView的所有行用于判断查询结束后是否有可见行
                    bool temp = true;//创建一个临时的布尔值用储存判断状态
                    for (int i = 0; i <= times; i++)
                    {
                        //对DataGridView控件的所有行进行判断，若存在可见项，将temp值置为false
                        if (StudentData.Rows[i].Visible != false)
                        {
                            temp = false;
                        }

                    }
                    if (temp != false)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        for (int j = 0; j <= times; j++)
                        {
                            StudentData.Rows[j].Visible = true;//将DataGridView的所有行的可见性还原
                        }
                        MessageBox.Show("查询模式已退出");
                        search = -1;
                        return;//这里没有清除姓名文本框的文本是为了让用户检查是否有输入错误(其实就是懒得写)
                    }

                    NameTextbox.Text = "";//同样为了减少操作，将姓名文本框的文本清除
                }
            }
        }

        private void 按性别ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可查询");
                return;
            }
            else
            {
                MessageBox.Show("您已进入查询模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                search = 2;//search等于2表示按性别查询模式
                if (GenderTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于查询的性别");
                    MessageBox.Show("查询模式已退出");
                    search = -1;//查询模式退出时将search的值还原为-1
                    return;
                }
                else
                {
                    //这一层for循环才是整个的查询操作，前面的if条件句均为是否可进行查询的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的姓名在数据列表中搜索信息，相关信息设置为可见，无关信息设置为不可见
                        if (GenderTextbox.Text == StudentData.Rows[i].Cells[search].Value.ToString())
                        {
                            StudentData.Rows[i].Visible = true;//搜索到的相关项设置为可见
                        }
                        else
                        {
                            StudentData.Rows[i].Visible = false;//无关项设置为不可见
                        }
                    }

                    //查询操作结束后再用一层for循环遍历DataGridView的所有行用于判断查询结束后是否有可见行
                    bool temp = true;//创建一个临时的布尔值用储存判断状态
                    for (int i = 0; i <= times; i++)
                    {
                        //对DataGridView控件的所有行进行判断，若存在可见项，将temp值置为false
                        if (StudentData.Rows[i].Visible != false)
                        {
                            temp = false;
                        }

                    }
                    if (temp != false)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        for (int j = 0; j <= times; j++)
                        {
                            StudentData.Rows[j].Visible = true;//将DataGridView的所有行的可见性还原
                        }
                        MessageBox.Show("查询模式已退出");
                        search = -1;
                        return;
                    }

                    GenderTextbox.Text = "";//同样为了减少操作，将性别文本框的文本清除
                }
            }
        }

        private void 按班号ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可查询");
                return;
            }
            else
            {
                if (CNumTextbox.Text.Length < 3)
                {
                    MessageBox.Show("班号输入不合法，应为3-4位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show("您已进入查询模式", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                search = 4;//search等于4表示按班号查询模式
                if (CNumTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于查询的班号");
                    MessageBox.Show("查询模式已退出");
                    search = -1;//查询模式退出时将search的值还原为-1
                    return;
                }
                else
                {
                    //这一层for循环才是整个的查询操作，前面的if条件句均为是否可进行查询的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的姓名在数据列表中搜索信息，相关信息设置为可见，无关信息设置为不可见
                        if (CNumTextbox.Text == StudentData.Rows[i].Cells[search].Value.ToString())
                        {
                            StudentData.Rows[i].Visible = true;//搜索到的相关项设置为可见
                        }
                        else
                        {
                            StudentData.Rows[i].Visible = false;//无关项设置为不可见
                        }
                    }

                    //查询操作结束后再用一层for循环遍历DataGridView的所有行用于判断查询结束后是否有可见行
                    bool temp = true;//创建一个临时的布尔值用储存判断状态
                    for (int i = 0; i <= times; i++)
                    {
                        //对DataGridView控件的所有行进行判断，若存在可见项，将temp值置为false
                        if (StudentData.Rows[i].Visible != false)
                        {
                            temp = false;
                        }

                    }
                    if (temp != false)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        for (int j = 0; j <= times; j++)
                        {
                            StudentData.Rows[j].Visible = true;//将DataGridView的所有行的可见性还原
                        }
                        MessageBox.Show("查询模式已退出");
                        search = -1;
                        return;//这里没有清除班号文本框的文本是为了让用户检查是否有输入错误(其实就是懒得写)
                    }

                    CNumTextbox.Text = "";//同样为了减少操作，将班号文本框的文本清除
                }
            }
        }

        private void 修改ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (StudentData.SelectedCells.Count == 0)//判断是否有选中行
            {
                MessageBox.Show("未选中需要修改数据的行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                int Select = StudentData.CurrentRow.Index;//声明变量用于保存选中行的行号
                int TrCount = 0;//声明变量用于记录文本框空文本的数量并赋初值为0
                //if条件句用于判断，若对应文本框的文本为空，则保持原有数据不变，若非空，则修改对应数据
                if (SNumTextbox.Text != "")
                {
                    if (SNumTextbox.Text.Length < 10)//将合法性判定改至对应的修改位置，避免出现修改其他信息时仍需输入学号班号
                    {
                        MessageBox.Show("学号输入不合法，学号应为10位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    StudentData.Rows[Select].Cells[0].Value = SNumTextbox.Text;//修改dataGridView控件的选中行学号列的单元格的元素
                    TrCount++;
                }
                if (NameTextbox.Text != "")
                {
                    StudentData.Rows[Select].Cells[1].Value = NameTextbox.Text;//修改dataGridView控件的选中行姓名列的单元格的元素
                    TrCount++;
                }
                if (GenderTextbox.Text != "")
                {
                    StudentData.Rows[Select].Cells[2].Value = GenderTextbox.Text;//修改dataGridView控件的选中行性别列的单元格的元素
                    TrCount++;
                }
                if (AgeTextbox.Text != "")
                {
                    StudentData.Rows[Select].Cells[3].Value = AgeTextbox.Text;//修改dataGridView控件的选中行年龄列的单元格的元素
                    TrCount++;
                }
                if (CNumTextbox.Text != "")
                {
                    if (CNumTextbox.Text.Length < 3)//将合法性判定改至对应的修改位置，避免出现修改其他信息时仍需输入学号班号
                    {
                        MessageBox.Show("班号输入不合法，应为3-4位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    StudentData.Rows[Select].Cells[4].Value = CNumTextbox.Text;//修改dataGridView控件的选中行班号列的单元格的元素
                    TrCount++;
                }
                if (SpecialityTextbox.Text != "")
                {
                    StudentData.Rows[Select].Cells[5].Value = SpecialityTextbox.Text;//修改dataGridView控件的选中行专业列的单元格的元素
                    TrCount++;
                }
                if (DepartmentTextbox.Text != "")
                {
                    StudentData.Rows[Select].Cells[6].Value = DepartmentTextbox.Text;//修改dataGridView控件的选中行系别列的单元格的元素
                    TrCount++;
                }
                if (TrCount == 0)
                {
                    MessageBox.Show("没有用于修改的数据，请确定是否需要修改数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    //信息录入后将所有文本框的文本清空，减少手动操作
                    SNumTextbox.Text = "";
                    NameTextbox.Text = "";
                    GenderTextbox.Text = "";
                    AgeTextbox.Text = "";
                    CNumTextbox.Text = "";
                    SpecialityTextbox.Text = "";
                    DepartmentTextbox.Text = "";
                }
            }
        }

        private void 按性别ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可用于统计");
                return;
            }
            else
            {
                if (GenderTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于统计的性别");
                    return;
                }
                else
                {
                    double CountNum = 0;//声明变量CountNum用于统计数据出现次数，赋初值为0
                    //这一层for循环才是整个的统计操作，前面的if条件句均为是否可进行统计的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的性别在数据列表中搜索数据，相同则计数器+1
                        if (GenderTextbox.Text == StudentData.Rows[i].Cells[2].Value.ToString())
                        {
                            CountNum++;//搜索到数据相同则计数器+1
                        }
                    }
                    //搜索操作结束后对计数器进行判断
                    if (CountNum == 0)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        double percent = CountNum / (times + 1) * 100.00;
                        MessageBox.Show("性别:" + GenderTextbox.Text + "\n" + "统计结果:占比" + percent + "%\n" + "人数:" + CountNum);
                    }
                }
            }
        }

        private void 按班号ToolStripMenuItem1_Click(object sender, EventArgs e)
        {            
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可用于统计");
                return;
            }
            else
            {
                if (CNumTextbox.Text.Length < 3)
                {
                    MessageBox.Show("班号输入不合法，应为3-4位数字!", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (CNumTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于统计的班号");
                    return;
                }
                else
                {
                    double CountNum = 0;//声明变量CountNum用于统计数据出现次数，赋初值为0
                    //这一层for循环才是整个的统计操作，前面的if条件句均为是否可进行统计的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的班号在数据列表中搜索数据，相同则计数器+1
                        if (CNumTextbox.Text == StudentData.Rows[i].Cells[4].Value.ToString())
                        {
                            CountNum++;//搜索到数据相同则计数器+1
                        }
                    }
                    //搜索操作结束后对计数器进行判断
                    if (CountNum == 0)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        double percent = CountNum / (times + 1) * 100.00;
                        MessageBox.Show("班号:" + CNumTextbox.Text + "\n" + "统计结果:占比" + percent + "%\n" + "人数:" + CountNum);
                    }
                }
            }
        }

        private void 按年龄ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可用于统计");
                return;
            }
            else
            {
                if (AgeTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于统计的年龄");
                    return;
                }
                else
                {
                    double CountNum = 0;//声明变量CountNum用于统计数据出现次数，赋初值为0
                    //这一层for循环才是整个的统计操作，前面的if条件句均为是否可进行统计的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的年龄在数据列表中搜索数据，相同则计数器+1
                        if (AgeTextbox.Text == StudentData.Rows[i].Cells[3].Value.ToString())
                        {
                            CountNum++;//搜索到数据相同则计数器+1
                        }
                    }
                    //搜索操作结束后对计数器进行判断
                    if (CountNum == 0)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        double percent = CountNum / (times + 1) * 100.00;
                        MessageBox.Show("年龄:" + AgeTextbox.Text + "\n" + "统计结果:占比" + percent + "%\n" + "人数:" + CountNum);
                    }
                }
            }
        }

        private void 按系别ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (times < 0)
            {
                MessageBox.Show("未录入学生信息，无信息可用于统计");
                return;
            }
            else
            {
                if (DepartmentTextbox.Text == "")
                {
                    MessageBox.Show("未输入用于统计的系别");
                    return;
                }
                else
                {
                    double CountNum = 0;//声明变量CountNum用于统计数据出现次数，赋初值为0
                    //这一层for循环才是整个的统计操作，前面的if条件句均为是否可进行统计的判断
                    for (int i = 0; i <= times; i++)//遍历DataGridView的所有行
                    {
                        //依据文本框中的系别在数据列表中搜索数据，相同则计数器+1
                        if (DepartmentTextbox.Text == StudentData.Rows[i].Cells[6].Value.ToString())
                        {
                            CountNum++;//搜索到数据相同则计数器+1
                        }
                    }
                    //搜索操作结束后对计数器进行判断
                    if (CountNum == 0)
                    {
                        MessageBox.Show("未找到有关数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        double percent = CountNum / (times + 1) * 100.00;
                        MessageBox.Show("系别:" + DepartmentTextbox.Text + "\n" + "统计结果:占比" + percent + "%\n" + "人数:" + CountNum);
                    }
                }
            }
        }

        private void 删除ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (StudentData.SelectedCells.Count == 0)//判断是否有选中行
            {
                MessageBox.Show("未选中需要删除的行", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                int Select = StudentData.CurrentRow.Index;//声明变量用于保存选中行的行号

                //删除选中行,但行索引没变,例如有行索引为0、1、2的三行,删除了行索引为1的行,其余两行仍为0、2而非0、1
                //这样会导致进行查询操作时遍历行索引时出现null值而报错,故废弃
                //StudentData.Rows.Remove(StudentData.Rows[Select]);

                for (int i = Select; i < times; i++)//从选中行开始遍历,将下一行的值赋给上一行
                {
                    StudentData.Rows[i].Cells[0].Value = StudentData.Rows[i + 1].Cells[0].Value;
                    StudentData.Rows[i].Cells[1].Value = StudentData.Rows[i + 1].Cells[1].Value;
                    StudentData.Rows[i].Cells[2].Value = StudentData.Rows[i + 1].Cells[2].Value;
                    StudentData.Rows[i].Cells[3].Value = StudentData.Rows[i + 1].Cells[3].Value;
                    StudentData.Rows[i].Cells[4].Value = StudentData.Rows[i + 1].Cells[4].Value;
                    StudentData.Rows[i].Cells[5].Value = StudentData.Rows[i + 1].Cells[5].Value;
                    StudentData.Rows[i].Cells[6].Value = StudentData.Rows[i + 1].Cells[6].Value;
                }
                StudentData.Rows.Remove(StudentData.Rows[times]);//遍历赋值操作结束后删除最末一行
                times--;//并将总行数减一
                MessageBox.Show("数据已删除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void 退出查询模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (search != -1)
            {
                for (int i = 0; i <= times; i++)
                {
                    StudentData.Rows[i].Visible = true;//将DataGridView的所有行的可见性还原
                }
                MessageBox.Show("查询模式已退出");
                search = -1;
            }
            else
            {
                MessageBox.Show("当前未处于查询模式", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 保存SToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DGVToExcel dGVToExcel = new DGVToExcel();
            dGVToExcel.OutputExcel(StudentData);
        }
    }

    public class DGVToExcel//新建一个类用于实现DataGridView控件的数据到Excel表格的转化
    {
        public Excel.Application DtoE = null;

        public void OutputExcel(DataGridView dataGridView)//用于保存excel文件的函数
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("无可用于保存的数据/空数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string FilePath = "";//声明变量用于储存文件路径
            SaveFileDialog dialog = new SaveFileDialog();//提示用户选择保存文件的位置
            dialog.Title = "保存数据到Excel文件";//提示框标题
            dialog.Filter = "Excel文件(*.xlsx)|*.xlsx";
            dialog.FilterIndex = 1;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FilePath = dialog.FileName;//如果用户选择了文件保存的路径，返回路径给变量FilePath
            }
            else
            {
                MessageBox.Show("未选择保存路径", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;//否则退出保存
            }

            //这个位置本来有一步将dataGridView转化为dataTable以过滤隐藏列(即控件创建时的首列*列)
            //但由于控件首列已经被禁用，所以省去
            //测试时发现多了一列空列处于列的尾部，可能是因为没有将dataGridView转化为dataTable导致

            long RowNum = dataGridView.Rows.Count;//声明变量以储存数据的行数
            long ColNum = dataGridView.Columns.Count;//声明变量以储存数据的列数
            Excel.Application DtoE = new Excel.Application();
            DtoE.DisplayAlerts = false;//隐藏更改提示
            DtoE.Visible = false;//默认不可见

            Excel.Workbooks workbooks = DtoE.Workbooks;
            Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//获取sheet1

            try
            {
                string[,] datas = new string[RowNum + 1, ColNum + 1];
                for (int c = 0; c < ColNum; c++)//写入数据
                {
                    datas[0, c] = dataGridView.Columns[c].HeaderText;//参考的资料中此处用的是Caption,但是发现无此标识符，故用同样作用的标识符HeaderText
                }
                Excel.Range range = DtoE.Range[worksheet.Cells[1, 1], worksheet.Cells[1, ColNum]];
                range.Interior.ColorIndex = 15;//设置标题行背景为灰色
                range.Font.Bold = true;//设置标题行文本为黑体
                range.Font.Size = 10;//设置标题行文本字号大小

                for (int r = 0; r < RowNum; r++)
                {
                    for (int c = 0; c < ColNum; c++)
                    {
                        datas[r+1,c] = dataGridView.Rows[r].Cells[c].Value.ToString();
                    }
                    System.Windows.Forms.Application.DoEvents();//添加进度条
                }
                Excel.Range FR = DtoE.Range[worksheet.Cells[1, 1], worksheet.Cells[RowNum + 1, ColNum + 1]];
                FR.Value2 = datas;

                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应
                DtoE.Visible = false;

                range = DtoE.Range[worksheet.Cells[1, 1], worksheet.Cells[RowNum + 1, ColNum + 1]];
                range.Font.Size = 9;
                range.RowHeight = 14.25;
                range.Borders.LineStyle = 1;
                range.HorizontalAlignment = 1;
                workbook.Saved = true;
                workbook.SaveCopyAs(FilePath);
            }
            catch(Exception error)
            {
                MessageBox.Show("保存文件异常"+error.Message,"错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

            DtoE.Workbooks.Close();
            DtoE.Workbooks.Application.Quit();
            DtoE.Application.Quit();
            GC.Collect();//回收内存
            MessageBox.Show("文件保存成功");
        }

    }
}
