package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Image;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.geom.RoundRectangle2D;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DecimalFormat;
import java.util.concurrent.CountDownLatch;

import javax.swing.ImageIcon;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.formdev.flatlaf.FlatIntelliJLaf;
import com.formdev.flatlaf.FlatLaf;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;

public class SplashAnimation extends JFrame {
	private static final long serialVersionUID = 3356025952135839862L;
	private JPanel contentPane;
	private JProgressBar jProgressBarTelaSplash;
	private JLabel jLabelMostraProgresso;
	private Image img_logo = new ImageIcon(SplashAnimation.class.getResource("/resources/Logo_HPE_menor.png"))
			.getImage().getScaledInstance(292, 73, Image.SCALE_SMOOTH);
	private boolean flagNecessarioFazerDownloading = false;
	private boolean flagLoading = false;
	private boolean flagError = false;
	private String userHome;
	private String fileName;
	private String pathRestante;
	private boolean[] dialogShown = { false }; // Array para armazenar a variável
	private CountDownLatch latch = new CountDownLatch(1);
	private File pathfileWord;
	private int finalOpcao;
	private File configDirectory = new File(System.getProperty("user.home") + "/ConversorXLSX-PDF");
	private final long originalSize = 30652;

	public SplashAnimation() {
		setUndecorated(true);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(0, 0, 382, 122);
		setLocationRelativeTo(null);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBackground(Color.WHITE);
		panel.setBounds(0, 0, 382, 122);
		contentPane.add(panel);
		panel.setLayout(null);

		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setIcon(new ImageIcon(img_logo));
		lblNewLabel.setBounds(0, 0, 382, 81);
		panel.add(lblNewLabel);

		jProgressBarTelaSplash = new JProgressBar();
		jProgressBarTelaSplash.setBounds(70, 72, 242, 14);
		panel.add(jProgressBarTelaSplash);

		jLabelMostraProgresso = new JLabel("");
		jLabelMostraProgresso.setFont(new Font("Dialog", Font.BOLD, 10));
		jLabelMostraProgresso.setHorizontalAlignment(SwingConstants.CENTER);
		jLabelMostraProgresso.setBounds(-1, 84, 382, 30);
		panel.add(jLabelMostraProgresso);

		TelaSplashCondominio();
	}

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					FlatLaf.registerCustomDefaultsSource("com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator");
					FlatMacDarkLaf.setup();
					SplashAnimation frame = new SplashAnimation();
					frame.setUndecorated(true);
					frame.setShape(new RoundRectangle2D.Double(0, 0, 382, 122, 20, 20));
					frame.setVisible(true);
					frame.setVisible(true);
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, "Erro ao abrir a janela: " + e);
				}
			}
		});
	}

	public void TelaSplashCondominio() {
		boolean[] opcaoSelecionada = { false };
		boolean[] arquivoValidado = { false };
		boolean[] configLinhaDiretorioImportado = { false };
		boolean[] configLinhaDiretorioRaiz = { false };
		boolean[] avisoflag = {true};
		
		new Thread() {
			public void run() {
				userHome = System.getProperty("user.home");
				fileName = "modelo laudo.docx";
				pathRestante = "/ConversorXLSX-PDF/data";

				for (int i = 0; i <= 100; i++) {
					try {
						sleep(30);
						jProgressBarTelaSplash.setValue(i);

						// 20%
						if (jProgressBarTelaSplash.getValue() <= 20) {
							jLabelMostraProgresso.setText("Verificando arquivos e pastas dependentes...");
							File pathExists = new File(userHome + pathRestante + "/" + fileName);
							File configFile = new File(configDirectory, "config.ini");
							// JOptionPane.showMessageDialog(null, pathExists.exists());

							if (configFile.exists() == false) {
								ConfigManager.createConfigFileIfNotExists();
								configLinhaDiretorioImportado[0] = ConfigManager.doesLineExist(3);
								configLinhaDiretorioRaiz[0] = ConfigManager.doesLineExist(4);
							} else {
								configLinhaDiretorioImportado[0] = ConfigManager.doesLineExist(3);
								configLinhaDiretorioRaiz[0] = ConfigManager.doesLineExist(4);
							}
							if (pathExists.exists() && configDirectory.exists() && configLinhaDiretorioRaiz[0]) {
								pathfileWord = pathExists;
								flagLoading = true;
							} else if (!pathExists.exists() || !configLinhaDiretorioRaiz[0]) {
								flagNecessarioFazerDownloading = true;
							}
						} else if (jProgressBarTelaSplash.getValue() <= 50 && flagNecessarioFazerDownloading == true) {
							// Selecionar uma opcao e pausa a thread para nao carregar a barra sozinha...
							if (opcaoSelecionada[0] == false) {
								do {
									SwingUtilities.invokeLater(() -> {
										String mensagem = "Deseja baixar o \"modelo laudo.docx\" via meu repositório GitHub?\n"
												+ "Obs.: se for baixar direto do git somente irá funcionar se o proxy não estiver bloqueando a requisição (URL de download) ou então desativando o proxy em opções de internet.\n\n"
												+ "Caso precise liberar a requisição bloqueada, basta pedir para equipe da infraestrutura a liberação dessa URL:\n"
												+ "https://github.com/GustavoBorges13/Conversor_XLSX-PDF/raw/main/data/modelo%20laudo.docx\n\n"
												+ "Sobre o arquivo a ser baixado, se trata de uma template personalizada que eu fiz onde existem campos com tags para facilitar no momento de transpor os dados das planilhas (Excel -> .xlsx) para a template (word -> .docx).\n\n"
												+ "Ou se você já estiver com arquivo baixado no computador ou no servidor, basta clicar no botão \"Procurar o arquivo no computador (local)\"";

										int opcao = mostrarDialogoPersonalizado(
												"Arquivo de modelo laudo.docx não foi encontrado!", mensagem,
												new String[] { "Baixar o modelo (online)",
														"Procurar o arquivo no computador (local)", "Cancelar" });
										switch (opcao) {
										case 0:
											opcaoSelecionada[0] = true;
											finalOpcao = 0;
											break;
										case 1:
											opcaoSelecionada[0] = true;
											finalOpcao = 1;

											// File direcionado = new File("\\\\fscatorg01\\...");
											// Defina o diretório inicial para a área de trabalho
											String userHome = System.getProperty("user.home");

											JFileChooser fc = new JFileChooser();
											fc.setCurrentDirectory(new File(userHome)); // Local padrao para abrir a
																						// janela
											fc.setPreferredSize(new Dimension(700, 400));

											// Restringir o usuário para selecionar arquivos de todos os tipos
											fc.setAcceptAllFileFilterUsed(false);

											// Coloca um titulo para a janela de dialogo
											fc.setDialogTitle("Selecione um arquivo .docx");

											// Habilita para o user escolher apenas arquivos do tipo: .docx
											FileNameExtensionFilter restrict = new FileNameExtensionFilter(
													"Somente arquivos .docx", "docx");
											fc.addChoosableFileFilter(restrict);
											fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

											int resultado = fc.showOpenDialog(null);

											// Verifica se um arquivo foi selecionado
											if (resultado == JFileChooser.CANCEL_OPTION) {
												System.exit(0);
											} else {
												pathfileWord = fc.getSelectedFile();

												// Verifica o tamanho do arquivo em kilobytes
												long fileSizeInBytes = pathfileWord.length();
												//long fileSizeInKB = fileSizeInBytes / 1024;

												// Verifica se o tamanho do arquivo está entre 27KB e 40KB
												// boolean tamanhoValido = fileSizeInKB > 27 && fileSizeInKB < 40;
												boolean tamanhoValido = fileSizeInBytes == originalSize; // 29.9 KB em
																											// bytes

												// Verifica se o nome do arquivo é "modelo laudo.docx"
												// boolean nomeValido = pathfileWord.getName().equals("modelo
												// laudo.docx");

												// Se ambos tamanho e nome são válidos
												// nomeValido removido..
												if (tamanhoValido) {
													/*
													 * JOptionPane.showMessageDialog(null,
													 * "Arquivo válido: tamanho está entre 27KB e 40KB, e o nome é \"modelo laudo.docx\"."
													 * );
													 */
													arquivoValidado[0] = true;

													// Define o diretório de destino
													File destino = new File(System.getProperty("user.home")
															+ "\\ConversorXLSX-PDF\\data");

													// Cria o diretório de destino se ainda não existir
													destino.mkdirs();

													// Cria o caminho de destino para o arquivo .docx
													File destinoArquivo = new File(destino, pathfileWord.getName());

													try {
														// Copia o arquivo para o diretório de destino
														Files.copy(pathfileWord.toPath(), destinoArquivo.toPath(),
																StandardCopyOption.REPLACE_EXISTING);

														// Verifica se o arquivo config.ini já existe, se não, cria
														File configFile = new File(configDirectory, "config.ini");
														if (!configFile.exists()) {
															configFile.createNewFile();
														}

														// Salva o caminho raiz do arquivo .docx do servidor ou da
														// pasta..
														ConfigManager.setConfigLine(3,
																pathfileWord.toPath().toString());

														// Salva o caminho do arquivo .docx no arquivo de configuração
														ConfigManager.setConfigLine(4, destino.getAbsolutePath());
													} catch (IOException e) {
														e.printStackTrace();
														JOptionPane.showMessageDialog(null,
																"Erro ao copiar o arquivo para o diretório de destino.","Erro", JOptionPane.ERROR_MESSAGE);
														System.exit(0);
													}
												} else {
													DecimalFormat decimalFormat = new DecimalFormat("0.0");
													JOptionPane.showMessageDialog(null,
															"O tamanho do arquivo de template do laudo não é compátivel. Por favor, baixe a template original via github.\n"
																	+ "Tamanho da template original: "
																	+ decimalFormat.format(originalSize / 1024f)
																	+ "KB (" + originalSize + " bytes)"
																	+ "\nTamanho da template selecionada: "
																	+ decimalFormat.format(fileSizeInBytes / 1024f)
																	+ "KB (" + fileSizeInBytes + " bytes)",
															"Aviso", JOptionPane.WARNING_MESSAGE);
													System.exit(0);
												}
											}

											break;

										case 2:
											JOptionPane.showMessageDialog(null,
													"Você não selecionou nenhum arquivo, o programa será fechado.","Erro", JOptionPane.ERROR_MESSAGE);
											System.exit(0);
										}

										// Marcar que o diálogo foi exibido
										dialogShown[0] = true;

										// Sinalizar que a thread pode continuar
										latch.countDown();
									});
									try {
										// Pausar a thread até que o diálogo seja fechado
										latch.await();
									} catch (InterruptedException e) {
										JOptionPane.showMessageDialog(null,
												"Erro Thread interrompida: " + e.getMessage(),"Erro", JOptionPane.ERROR_MESSAGE);
									}
								} while (!dialogShown[0]);

								// Continua executando a thread... barra de progresso... sem mostrar a
								// joptionpane novamente.
							} else {
								if (finalOpcao == 0 && opcaoSelecionada[0] == true && arquivoValidado[0] == false) {
									jLabelMostraProgresso.setText("Fazendo download do modelo de laudo técnico...");
									downloadFile();
								} else if (finalOpcao == 1 && opcaoSelecionada[0] == true
										&& arquivoValidado[0] == true) {
									jLabelMostraProgresso.setText("Criando/atualizando o arquivo config.ini...");
								}
							}
						} else if (jProgressBarTelaSplash.getValue() <= 60 && flagLoading == true) {
							jLabelMostraProgresso.setText("Carregando o modelo de laudo técnico...");
							if(!(pathfileWord.length() == originalSize)&&avisoflag[0]==true) {
								DecimalFormat decimalFormat = new DecimalFormat("0.0");
								long fileSizeInBytes = pathfileWord.length();
								JOptionPane.showMessageDialog(null, "O tamanho atual do arquivo de template instalado na máquina não batem com os valores programados no sistema, recomenda-se fortemente a exclusão manual da template em uso."
															+"\nTemplate localização: "+pathfileWord.getAbsolutePath()
															+ "\nTamanho da template original: "
															+ decimalFormat.format(originalSize / 1024f)
															+ "KB (" + originalSize + " bytes)"
															+ "\nTamanho da template selecionada: "
															+ decimalFormat.format(fileSizeInBytes / 1024f)
															+ "KB (" + fileSizeInBytes + " bytes)",
															"Aviso", JOptionPane.WARNING_MESSAGE);
								avisoflag[0]=false;
							}
						} else if (jProgressBarTelaSplash.getValue() <= 80) {
							jLabelMostraProgresso.setText("Ajustando a UI..");
							// 100%
						} else if (jProgressBarTelaSplash.getValue() == 100 && flagError == false) {
							jLabelMostraProgresso.setText("Conectando com o sistema...");
							Principal tela = new Principal();
							tela.setVisible(true);
							tela.requestFocus();
							SplashAnimation.this.dispose();
						}

					} catch (Exception e) {
						flagError = false;
						e.printStackTrace(); // Adicione esta linha para imprimir o rastreamento da exceção
						JOptionPane.showMessageDialog(null, "Erro crítico (1): " + e,"Erro", JOptionPane.ERROR_MESSAGE);
					}
				}
			}
		}.start();
	}

	private static int mostrarDialogoPersonalizado(String titulo, String mensagem, String[] opcoes) {
		JTextArea textArea = new JTextArea(mensagem);
		textArea.setEditable(false);
		textArea.setLineWrap(true);
		textArea.setWrapStyleWord(true);

		JScrollPane scrollPane = new JScrollPane(textArea);
		scrollPane.setPreferredSize(new Dimension(400, 200));

		JOptionPane optionPane = new JOptionPane(scrollPane, JOptionPane.PLAIN_MESSAGE, JOptionPane.DEFAULT_OPTION,
				null, opcoes, opcoes[0]);

		JDialog dialog = optionPane.createDialog(titulo);

		// Adicione um WindowListener para detectar o fechamento da janela
		dialog.addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent e) {
				// Se o usuário fechar a janela, saia do programa
				JOptionPane.showMessageDialog(null, "Você escolheu sair do programa!", "Aviso",
						JOptionPane.WARNING_MESSAGE);
				System.exit(0);
			}
		});

		dialog.setVisible(true);

		// Retorna o valor escolhido pelo usuário
		Object selectedValue = optionPane.getValue();
		if (selectedValue == null)
			return JOptionPane.CLOSED_OPTION;

		for (int counter = 0, maxCounter = opcoes.length; counter < maxCounter; counter++) {
			if (opcoes[counter].equals(selectedValue))
				return counter;
		}

		// Se chegou aqui, algo deu errado
		return JOptionPane.CLOSED_OPTION;
	}

	public void downloadFile() {
		LeitorURL baixarArquivo = new LeitorURL();

		// URL que aponta para o arquivo a ser baixado
		String sUrl = "https://github.com/GustavoBorges13/Conversor_XLSX-PDF/raw/main/data/modelo%20laudo.docx";

		URL url = null;
		try {
			url = new URL(sUrl);
		} catch (MalformedURLException e) {
			flagError = true;
			JOptionPane.showMessageDialog(null,
					"Erro critico (2): " + e + "\nPossivel problema com proxy ou offline (sem acesso a internet).",
					"Erro", JOptionPane.ERROR_MESSAGE);
			System.exit(0);
		}

		// Local onde será baixado
		File folder = new File(userHome + pathRestante);
		folder.mkdirs();
		File file = new File(folder, fileName);
		baixarArquivo.copyURLToFile(url, file);

		try {

			// Verifica se o arquivo config.ini já existe, se não, cria
			File configFile = new File(configDirectory, "config.ini");
			if (!configFile.exists()) {
				configFile.createNewFile();
			}

			// Salva o caminho raiz do arquivo .docx do servidor ou da pasta..
			ConfigManager.setConfigLine(3, "N/A");

			// Salva o caminho do arquivo .docx no arquivo de configuração
			ConfigManager.setConfigLine(4, file.getAbsolutePath());

		} catch (IOException e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(null, "Erro ao copiar o arquivo para o diretório de destino.", "Erro",
					JOptionPane.ERROR_MESSAGE);
			System.exit(0);
		}
	}

	@Deprecated
	public void atualizarFile() {
		userHome = System.getProperty("user.home");
		fileName = "modelo laudo.docx";
		pathRestante = "/Documents/ConversorXLSX-PDF/data";

		// URL que aponta para o arquivo a ser baixado
		String sUrl = "https://github.com/GustavoBorges13/Conversor_XLSX-PDF/raw/main/data/modelo%20laudo.docx";

		URL url = null;
		try {
			url = new URL(sUrl);
		} catch (MalformedURLException e) {
			flagError = true;
			JOptionPane.showMessageDialog(null,
					"Não foi possível encontrar os arquivos necessários para a atualização." + e, "Erro",
					JOptionPane.ERROR_MESSAGE);
			dispose();
		}

		// Local onde será baixado
		File folder = new File(userHome + pathRestante);
		folder.mkdirs();

		try {
			InputStream inputStream = url.openStream();
			Files.copy(inputStream, Paths.get(folder.getPath() + "\\" + fileName), StandardCopyOption.REPLACE_EXISTING);
		} catch (IOException e) {
			flagError = true;
			JOptionPane.showMessageDialog(null,
					"Erro critico (2): " + e + "\nPossivel erro de proxy ou offline (sem acesso a internet).", "Erro",
					JOptionPane.ERROR_MESSAGE);
			System.exit(0);
		}
	}
}
