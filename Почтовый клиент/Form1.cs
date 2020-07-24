using MailKit;
using MailKit.Net.Imap;
using MimeKit;

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Почтовый_клиент
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.login = textBox2.Text;
            Properties.Settings.Default.password = textBox3.Text;
            Properties.Settings.Default.Save();

            string from = textBox2.Text;
            string password = textBox3.Text;
            string to = textBox4.Text;
            string subject = textBox5.Text;
            string body = textBox6.Text;

            if (checkBox1.Checked)
            {
                string fileNameAttachment = textBox1.Text;
                MailMessage message = Send(from, to, password, subject, body, fileNameAttachment);
                ToSent(from, password, message);
            }
            else
            {
                MailMessage message = Send(from, to, password, subject, body);
                ToSent(from, password, message);
            }
        }

        private void ToSent(string login, string password, MailMessage message) 
        {
            ImapClient imap = new ImapClient();
            imap.Connect("imap.mail.ru", 993, true);
            imap.Authenticate(login, password);
            FolderNamespace folderNamespaceSent = new FolderNamespace('/', "Отправленные");
            IMailFolder folderSent = imap.GetFolder(folderNamespaceSent);
            folderSent.Open(FolderAccess.ReadOnly);
            MimeMessage mimeMessage = (MimeMessage)message;          
            folderSent.Append(mimeMessage);
        }

        private MailMessage Send(string from, string to, string password, string subject, string body)
        {
            MailMessage message = new MailMessage(from, to);
            // тема письма
            message.Subject = subject;
            // текст письма
            message.Body = body;
            //m.CC.Add(from); //Скрытая копия
            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            SmtpClient smtp = new SmtpClient("smtp.mail.ru", 25);
            // логин и пароль
            smtp.Credentials = new NetworkCredential(from, password);
            smtp.EnableSsl = true;
            smtp.Send(message);
            return message;
        }

        private MailMessage Send(string from, string to, string password, string subject, string body, string fileNameAttachment)
        {
            MailMessage message = new MailMessage(from, to);
            // тема письма
            message.Subject = subject;
            // текст письма
            message.Body = body;
            Attachment attachment = new Attachment(fileNameAttachment);
            message.Attachments.Add(attachment);
            //m.CC.Add(from); //Скрытая копия
            // адрес smtp-сервера и порт, с которого будем отправлять письмо
            SmtpClient smtp = new SmtpClient("smtp.mail.ru", 25);
            // логин и пароль
            smtp.Credentials = new NetworkCredential(from, password);
            smtp.EnableSsl = true;
            smtp.Send(message);
            return message;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Enabled = true;
                    textBox1.Text = openFileDialog1.FileName;
                }
            }
            else
            {
                textBox1.Enabled = false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Enabled = false;

            textBox2.Text = Properties.Settings.Default.login;
            textBox3.Text = Properties.Settings.Default.password;

            StringCollection patterns = Properties.Settings.Default.patterns;

            for (int i = 0; i < patterns.Count; i++)
            {
                comboBox1.Items.Add(patterns[i].Split('|')[0]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.login = textBox2.Text;
            Properties.Settings.Default.password = textBox3.Text;
            Properties.Settings.Default.Save();

            string from = textBox2.Text;
            string password = textBox3.Text;
            string subject = textBox9.Text;
            string fileNameAttachment = textBox1.Text;

            string[] rows = File.ReadAllLines(textBox7.Text);
            string[] header = rows[0].Split(';');

            for (int i = 1; i < rows.Length; i++)
            {
                string[] row = rows[i].Split(';');

                string body = textBox8.Text;

                for (int k = 0; k < header.Length; k++)
                {
                    body = body.Replace($"<{header[k]}>", row[k]);
                }

                string to = row[0];

                MailMessage message = Send(from, to, password, subject, body, fileNameAttachment);
                ToSent(from, password, message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string pattern = $"{textBox10.Text}|" +
                $"{textBox9.Text}|" +
                $"{textBox8.Text}";

            Properties.Settings.Default.patterns.Add(pattern);
            Properties.Settings.Default.Save();

            comboBox1.Items.Add(textBox10.Text);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] pattern = Properties.Settings.Default.patterns[comboBox1.SelectedIndex].Split('|');

            textBox9.Text = pattern[1];
            textBox8.Text = pattern[2];
        }
    }
}
