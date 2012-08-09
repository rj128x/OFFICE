using System.IO;

namespace Office.Shared
{	
	public delegate bool ProcessFileDelegate(string fileName);
	class FolderOperations
	{
		public ProcessFileDelegate processFile=null;
		public void calcFolder(string path) {
			DirectoryInfo di=new DirectoryInfo(path);
			System.IO.FileInfo[] files = null;
			System.IO.DirectoryInfo[] subDirs = null;

			files = di.GetFiles("*");

			if (files != null) {
				foreach (System.IO.FileInfo fi in files) {
					if (processFile != null) {
						processFile(fi.FullName);
					}
					System.Windows.Forms.Application.DoEvents();
				}
				subDirs = di.GetDirectories();

				foreach (System.IO.DirectoryInfo dirInfo in subDirs) {
					calcFolder(dirInfo.FullName);
				}
			}     
		}
	}
}
