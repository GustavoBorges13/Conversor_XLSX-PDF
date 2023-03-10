package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Image;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.SwingConstants;
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
	
	public SplashAnimation() {
		setUndecorated(true);
		TelaSplashCondominio();
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
		panel.setBounds(0, 0, 381, 122);
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
	}

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(new FlatIntelliJLaf());
					SplashAnimation frame = new SplashAnimation();
					frame.setVisible(true);
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, "Erro ao abrir a janela: "+e);
				}
			}
		});
	}

	public void TelaSplashCondominio() {
		new Thread() {
			public void run() {
				userHome = System.getProperty("user.home");
				fileName = "modelo laudo.docx";
				pathRestante = "/Documents/ConversorXLSX-PDF/data";
				File pathVerify = new File(userHome+"/Documents");
				if(pathVerify.exists() == false) {
					JOptionPane.showMessageDialog(null, "Nao existe");
					pathRestante =  "/Documentos/ConversorXLSX-PDF/data";
				}
				
				for (int i = 0; i < 101; i++) {
					try {
						sleep(30);
						jProgressBarTelaSplash.setValue(i);

						// 40%
						if (jProgressBarTelaSplash.getValue() <= 20) {
							jLabelMostraProgresso.setText("Verificando arquivos de integridade...");
							File pathExists = new File(userHome + pathRestante + "/" + fileName);

							if (pathExists.exists()) {
								flagLoading = true;
							} else {
								flagDownloading = true;
							}

						} else if (jProgressBarTelaSplash.getValue() <= 50 && flagLoading == true) {

							jLabelMostraProgresso.setText("Carregando o modelo de laudo t??cnico...");
							flagNecessarioAtualizar = true;
							// }else if(!pathExists.exists() && flagDownloading==false){

						} else if (jProgressBarTelaSplash.getValue() <= 50 && flagDownloading == true) {
							jLabelMostraProgresso.setText("Fazendo download do modelo de laudo t??cnico...");
							downloadFile();
							flagNecessarioAtualizar = false;
						} else if (jProgressBarTelaSplash.getValue() <= 60 && flagNecessarioAtualizar == true) {
							jLabelMostraProgresso.setText("Atualizando o modelo de laudo t??cnico...");
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
						JOptionPane.showMessageDialog(null, "Erro critico (1): "+e);
					}
				}
			}
		}.start();
	}

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
					"N??o foi poss??vel encontrar os arquivos necess??rios para o download." + e);
			dispose();
		}

		// Local onde ser?? baixado
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
					"N??o foi poss??vel encontrar os arquivos necess??rios para a atualiza????o." + e);
			dispose();
		}

		// Local onde ser?? baixado
		File folder = new File(userHome + pathRestante);
		folder.mkdirs();

		try {
			InputStream inputStream = url.openStream();
			Files.copy(inputStream, Paths.get(folder.getPath() + "\\" + fileName), StandardCopyOption.REPLACE_EXISTING);
		} catch (IOException e) {
			flagError = true;
			JOptionPane.showMessageDialog(null, "Erro critico (2): "+e+"\nPossivel erro de proxy ou offline (sem acesso a internet).");
			System.exit(0);
		}
	}
}
