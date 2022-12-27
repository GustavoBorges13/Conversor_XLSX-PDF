package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;

import javax.swing.JOptionPane;

@SuppressWarnings("unused")
public class ManipuladorDeArquivo {
	public void leitor(String path, String nomeDoArquivo) throws IOException {
		try (BufferedReader buffRead = new BufferedReader(new FileReader(path + nomeDoArquivo))) {
			String[] split;// Vetor de separacao de cada conteudo na linha

			// - - - Lendo o arquivo...
			String linha = buffRead.readLine(); // ler a primeira linha do arquivo
			try {
				while (linha != null) {
					/*
					 * split = linha.split(" - "); // Separa o nome do path
					 * 
					 * InterfaceJframe.fileNameB.add(split[0]); // Salva a leitura do nome do
					 * arquivo na lista de file names. InterfaceJframe.filePathB.add(split[1]); //
					 * Salva a leitura do path para na lista de paths dir.
					 */
					linha = buffRead.readLine(); // Linha p/ linha
				}
			} catch (ArrayIndexOutOfBoundsException e) {
				e.printStackTrace();
			}
			buffRead.close();
		} catch (IOException e) {
			System.out.println("Error: " + e.getMessage());
		}
	}

	public void escritor(String path, String nomeDoArquivo)
			throws IOException {
		File file = new File(path + nomeDoArquivo);
		file.getParentFile().mkdirs();
		FileWriter writer = new FileWriter(file);
		writer.close();
	}

	public void escritorv2(String path, String nomeDoArquivo, String fileName, String filePath) throws IOException {
		File file = new File(path + nomeDoArquivo);
		file.getParentFile().mkdirs();
		FileWriter writer = new FileWriter(file, true);

		try {
			BufferedWriter buffwriter = new BufferedWriter(writer);
			buffwriter.write(fileName + " - " + filePath + "\n");
			buffwriter.close();
		} catch (IOException ex) {
			System.out.println("File could not be created");
		}
	}
}
