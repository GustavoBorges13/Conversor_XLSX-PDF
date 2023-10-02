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
import java.util.concurrent.CountDownLatch;

import javax.swing.ImageIcon;
import javax.swing.JDialog;
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

import com.formdev.flatlaf.FlatIntelliJLaf;

public class SplashAnimation extends JFrame {
	private static final long serialVersionUID = 3356025952135839862L;
	private JPanel contentPane;
	private JProgressBar jProgressBarTelaSplash;
	private JLabel jLabelMostraProgresso;
	private Image img_logo = new ImageIcon(SplashAnimation.class.getResource("/resources/Logo_HPE_menor.png"))
			.getImage().getScaledInstance(292, 73, Image.SCALE_SMOOTH);

	private boolean flagNecessarioAtualizar = false;
	private boolean flagDownloading = false;
	private boolean flagLoading = false;
	private boolean flagError = false;
	private String userHome;
	private String fileName;
	private String pathRestante;
	private boolean[] dialogShown = { false }; // Array para armazenar a variável
	private CountDownLatch latch = new CountDownLatch(1);

	public SplashAnimation() {
		setUndecorated(true);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 381, 122);
		setLocationRelativeTo(null);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBackground(Color.WHITE);
		panel.setBounds(0, 0, 420, 122);
		contentPane.add(panel);
		panel.setLayout(null);

		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setIcon(new ImageIcon(img_logo));
		lblNewLabel.setBounds(0, 0, 381, 81);
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
					UIManager.setLookAndFeel(new FlatIntelliJLaf());
					SplashAnimation frame = new SplashAnimation();
					frame.setUndecorated(true);
					frame.setShape(new RoundRectangle2D.Double(0, 0, 420, 122, 25, 25));
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
		int opcao = 0;
		new Thread() {
			public void run() {
				userHome = System.getProperty("user.home");
				fileName = "modelo laudo.docx";
				pathRestante = "/Documents/ConversorXLSX-PDF/data";
				File pathVerify = new File(userHome + "/Documents");
				if (pathVerify.exists() == false) {
					JOptionPane.showMessageDialog(null,
							"Pasta \"Documents\" não existe. Portanto será criado uma pasta dentro de: " + userHome);
					pathRestante = "/Documentos/ConversorXLSX-PDF/data";
				}

				for (int i = 0; i <= 100; i++) {
					try {
						sleep(30);
						jProgressBarTelaSplash.setValue(i);

						// 40%
						if (jProgressBarTelaSplash.getValue() <= 20) {
							jLabelMostraProgresso.setText("Verificando arquivos e pastas dependentes...");
							File pathExists = new File(userHome + pathRestante + "/" + fileName);
							// JOptionPane.showMessageDialog(null, pathExists.exists());
							if (pathExists.exists()) {
								flagLoading = true;
							} else {
								flagDownloading = true;
							}
						} else if (jProgressBarTelaSplash.getValue() <= 50 && flagLoading == true) {
							jLabelMostraProgresso.setText("Carregando o modelo de laudo técnico...");
							flagNecessarioAtualizar = true;
						} else if (jProgressBarTelaSplash.getValue() <= 50 && flagDownloading == true) {
							// Selecionar uma opcao e pausa a thread para nao carregar a barra sozinha...
							if (opcaoSelecionada[0] == false) {
								do {
									SwingUtilities.invokeLater(() -> {
										String mensagem = "Deseja baixar o \"modelo laudo.docx\" via meu repositório GitHub?\n"
												+ "Obs.: baixar direto do git somente funciona se o proxy não estiver bloqueando ou desativando o proxy).\n\n"
												+ "Caso precise liberar a requisição bloqueada, basta pedir para equipe da infraestrutura a liberação dessa URL:\n"
												+ "https://github.com/GustavoBorges13/Conversor_XLSX-PDF/raw/main/data/modelo%20laudo.docx";

										int opcao = mostrarDialogoPersonalizado(
												"Arquivo de modelo laudo.docx não foi encontrado!", mensagem,
												new String[] { "Baixar o modelo (online)",
														"Procurar o arquivo no computador (local)", "Cancelar" });

										switch (opcao) {
										case 0:
											opcaoSelecionada[0] = true;
											opcao = 0;
											break;
										case 1:
											opcaoSelecionada[0] = true;
											opcao = 1;
											break;
										case 2:
											JOptionPane.showMessageDialog(null, "Você escolheu sair do programa!");
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
										e.printStackTrace();
									}
								} while (!dialogShown[0]);

								// Continua executando a thread... barra de progresso... sem mostrar a
								// joptionpane novamente.
							} else {
								if (opcao == 0 && opcaoSelecionada[0] == true) {
									jLabelMostraProgresso.setText("Fazendo download do modelo de laudo técnico...");
									downloadFile();
									flagNecessarioAtualizar = false;
								} else if (opcao == 1 && opcaoSelecionada[0] == true) {

								}
							}
						} else if (jProgressBarTelaSplash.getValue() <= 60 && flagNecessarioAtualizar == true) {
							jLabelMostraProgresso.setText("Atualizando o modelo de laudo técnico...");
							atualizarFile();
							// 70%
						} else if (jProgressBarTelaSplash.getValue() <= 80) {
							jLabelMostraProgresso.setText("Carregando banco de dados...");
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
						flagError = false;
						e.printStackTrace(); // Adicione esta linha para imprimir o rastreamento da exceção
						JOptionPane.showMessageDialog(null, "Erro crítico (1): " + e);
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

	        JOptionPane optionPane = new JOptionPane(scrollPane, JOptionPane.PLAIN_MESSAGE,
	                JOptionPane.DEFAULT_OPTION, null, opcoes, opcoes[0]);

	        JDialog dialog = optionPane.createDialog(titulo);

	        // Adicione um WindowListener para detectar o fechamento da janela
	        dialog.addWindowListener(new WindowAdapter() {
	            @Override
	            public void windowClosing(WindowEvent e) {
	                // Se o usuário fechar a janela, saia do programa
	            	JOptionPane.showMessageDialog(null, "Você escolheu sair do programa!");
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

	@Deprecated
	public void downloadFile() {
		LeitorURL baixarArquivo = new LeitorURL();

		// URL que aponta para o arquivo a ser baixado
		String sUrl = "https://github.com/GustavoBorges13/Conversor_XLSX-PDF/raw/main/data/modelo%20laudo.docx";

		URL url = null;
		try {
			url = new URL(sUrl);
		} catch (MalformedURLException e) {
			// TODO Auto-generated catch block
			flagError = false;
			JOptionPane.showMessageDialog(null,
					"Não foi possível encontrar os arquivos necessários para o download." + e);
			dispose();
		}

		// Local onde será baixado
		File folder = new File(userHome + pathRestante);
		folder.mkdirs();
		File file = new File(folder, fileName);
		baixarArquivo.copyURLToFile(url, file);
	}

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
					"Não foi possível encontrar os arquivos necessários para a atualização." + e);
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
					"Erro critico (2): " + e + "\nPossivel erro de proxy ou offline (sem acesso a internet).");
			System.exit(0);
		}
	}
}
