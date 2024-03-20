package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JOptionPane;

public class ConfigManager {
	private static final String CONFIG_FILE = System.getProperty("user.home")
			+ "\\ConversorXLSX-PDF\\config.ini";

	// Método para verificar se a linha existe no arquivo de configuração
	public static boolean doesLineExist(int lineNumber) {
		List<String> lines = getConfigFileLines();
		return lineNumber >= 0 && lineNumber < lines.size() && lines.get(lineNumber) != null;
	}

	// Método auxiliar para criar o arquivo de configuração se não existir
	public static void createConfigFileIfNotExists() {
		try {
			// Antes de salvar o arquivo config.ini
			File configDir = new File(System.getProperty("user.home") + "\\ConversorXLSX-PDF");
			configDir.mkdirs(); // Cria o diretório se ainda não existir

			File configFile = new File(CONFIG_FILE);
			if (!configFile.exists()) {
				configFile.createNewFile();

				// Cria um comentário na linha 1 para não substituir
				ConfigManager.setConfigLine(0, System.getProperty("user.home") + "\\ConversorXLSX-PDF\\config.ini");
				
				//Criando a linha 2 - paths backup
				String userHome = System.getProperty("user.home");
				String pathRestante = "\\ConversorXLSX-PDF\\data\\";
				File path= new File(userHome + pathRestante + "Backup");
				ConfigManager.setConfigLine(1, path.toPath().toString());
				
				//Criando a linha 3 - paths pdf generated
				path = new File(userHome + pathRestante + "Pdf generated");
				ConfigManager.setConfigLine(2, path.toPath().toString());
			}
		} catch (IOException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "Erro ao copiar o arquivo para o diretório de destino.");
			System.exit(0);
		}
	}

	// Método para obter o diretório do arquivo na linha especificada do arquivo de
	// configuração
	public static String getDirectoryFromConfigLine(int lineNumber) {
		List<String> lines = getConfigFileLines();

		// Verificar se a linha especificada existe
		if (lineNumber >= 0 && lineNumber < lines.size()) {
			String line = lines.get(lineNumber);

			// Verificar se a linha não está vazia
			if (line != null && !line.isEmpty()) {
				return line.trim();
			}
		}

		// Se a linha não existir ou estiver vazia, retornar null ou um diretório
		// padrão, conforme necessário
		return null;
	}

	public static void setConfigLine(int lineNumber, String content) {
		// Verifica se o arquivo de configuração existe, se não, cria
		File configFile = new File(System.getProperty("user.home") + "\\ConversorXLSX-PDF\\config.ini");
		if (!configFile.exists()) {
			try {
				configFile.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
				System.exit(0);
			}
		}

		// Lê as linhas atuais do arquivo
		List<String> lines = getConfigFileLines();

		// Preenche as linhas até o número desejado
		while (lines.size() <= lineNumber) {
			lines.add(""); // Adiciona linhas em branco se necessário
		}

		// Define o conteúdo na linha especificada
		lines.set(lineNumber, content);

		// Escreve as linhas atualizadas de volta no arquivo
		try (BufferedWriter writer = new BufferedWriter(new FileWriter(configFile))) {
			for (String line : lines) {
				writer.write(line);
				writer.newLine();
			}
		} catch (IOException e) {
			e.printStackTrace();
			System.exit(0);
		}
	}

	private static List<String> getConfigFileLines() {
		List<String> lines = new ArrayList<>();

		try (BufferedReader reader = new BufferedReader(new FileReader(CONFIG_FILE))) {
			String line;
			while ((line = reader.readLine()) != null) {
				lines.add(line);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		return lines;
	}

}
