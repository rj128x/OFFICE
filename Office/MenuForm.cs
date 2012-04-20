using System;
using System.Windows.Forms;

namespace Office
{
	public partial class MenuForm : Form
	{
		protected PDFForm blanksForm;
		protected ListsZnakForm listsZnakForm;
        protected ProtokolForm ocencaForm;
		  protected CreatePDFForm pdfForm;
		public MenuForm() {
			InitializeComponent();
		}

		private void btnBlanks_Click(object sender, EventArgs e) {
			if (blanksForm == null) {
				blanksForm = new PDFForm();
			}
			blanksForm.ShowDialog();
		}

		private void btnListsZnak_Click(object sender, EventArgs e) {
			
			if (listsZnakForm == null) {
				listsZnakForm = new ListsZnakForm();
			}
			listsZnakForm.ShowDialog();
		}

        private void btnOcenca_Click(object sender, EventArgs e)
        {
            if (ocencaForm == null)
            {
                ocencaForm = new ProtokolForm();
            }
            ocencaForm.ShowDialog();
        }

		  private void button1_Click(object sender, EventArgs e) {
			  if (pdfForm == null) {
				  pdfForm = new CreatePDFForm();
			  }
			  pdfForm.ShowDialog();
		  }
	}
}
