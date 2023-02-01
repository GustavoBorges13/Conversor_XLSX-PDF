package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import javax.swing.JOptionPane;

public class URLReader {
	public void copyURLToFile(URL url, File file) {
		try {
			InputStream input = url.openStream();
			if (file.exists()) {
				if (file.isDirectory())
					throw new IOException("Arquivo '" + file + "' é um diretório");

				if (!file.canWrite())
					throw new IOException("Arquivo '" + file + "' não pode ser escrito");
			} else {
				File parent = file.getParentFile();
				if ((parent != null) && (!parent.exists()) && (!parent.mkdirs())) {
					throw new IOException("Arquivo '" + file + "' não pode ser criado");
				}
			}

			FileOutputStream output = new FileOutputStream(file);

			byte[] buffer = new byte[4096];
			int n = 0;
			while (-1 != (n = input.read(buffer))) {
				output.write(buffer, 0, n);
			}

			input.close();
			output.close();
		} catch (IOException ioEx) {
			JOptionPane.showMessageDialog(null, "Não foi possível encontrar os arquivos necessários para o download."
					+ioEx);
			System.exit(0);
		}
	}
}