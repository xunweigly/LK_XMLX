using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UFIDA.U8.UAP.UI.Runtime.Model;
using System.IO;

namespace U8SOFT.XMRZ
{
    public partial class UserControl2 : PictureBox
    {
          private BusinessProxy businessProxy = null;
        private VoucherProxy voucherProxy = null;
        public UserControl2(BusinessProxy businessProxy, VoucherProxy voucherProxy)
        {
            InitializeComponent();
            this.businessProxy = businessProxy;
            this.voucherProxy = voucherProxy;
        }

        private void 粘贴ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            IDataObject iData = Clipboard.GetDataObject();
            if (iData.GetDataPresent(DataFormats.Bitmap))
            {
                this.Image = (Bitmap)iData.GetData(DataFormats.Bitmap);

            }

        }


        private void 清除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Image =  null;
        }

        private void 另存为_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            string pictureName;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                pictureName = saveFileDialog1.FileName;



                if (this.Image != null)
                {

                    ////********************照片另存*********************************

                    using (MemoryStream mem = new MemoryStream())
                    {

                        //这句很重要，不然不能正确保存图片或出错（关键就这一句）

                        Bitmap bmp = new Bitmap(this.Image);

                        //保存到内存

                        //bmp.Save(mem, pictureBox1.Image.RawFormat );

                        //保存到磁盘文件

                        bmp.Save(@pictureName, this.Image.RawFormat);

                        bmp.Dispose();



                        MessageBox.Show("图片另存成功！", "系统提示");

                    }

                    ////********************照片另存*********************************

                }



            }

        }
        
        
        //public UserControl2()
        //{
        //    InitializeComponent();
        //}
    }
}
