package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.KeyListener;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImageOp;
import java.awt.image.ConvolveOp;
import java.awt.image.Kernel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import javax.swing.DefaultComboBoxModel;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JEditorPane;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.Document;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.undo.UndoManager;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.formdev.flatlaf.FlatLaf;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;

public class GerarLaudoPDF extends JDialog {
	private static final long serialVersionUID = 4893492449132639712L;
	private JPanel contentPane;
	private static JTextField txtNomeTecnico;
	private static JTextField txtUsuarioRede;
	private static JTextField txtCentroCusto;
	static JEditorPane editorPaneConsideracoesTecnicas;
	static JTextArea textAreaAnalise;
	static int[] linhasSelecionadas;
	private static String consideracoesTecnicasTEMP;
	private DefaultStyledDocument doc;
	private JLabel remaningLabel = new JLabel("");;
	private int textAreaCaracteresLimit = 440;
	static boolean finalizado;
	private static JLabel lblNewLabel;
	private static JButton btnVisualizar;
	private static JButton btnAbrirLocal;
	static String pdfPATH;
	boolean flagCheckBox = false;
	private JButton btnAdd;
	private JButton btnDel;
	private String tempText;
	private int num = 0;
	public static List<String> linkExibicao = new ArrayList<String>();
	public static List<String> linkEndereco = new ArrayList<String>();
	@SuppressWarnings("rawtypes")
	private static JComboBox comboBoxTemplate;
	private static String textRef;
	private JCheckBox chckbxNewCheckBox;
	private JButton btnGerarArquivoPDF;
	private Image img_help = new ImageIcon(SplashAnimation.class.getResource("/resources/help-icon.png")).getImage()
			.getScaledInstance(20, 20, Image.SCALE_SMOOTH);
	private BufferedImage convolvedImage;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					FlatLaf.registerCustomDefaultsSource(
							"com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator");
					FlatMacDarkLaf.setup();
					GerarLaudoPDF frame = new GerarLaudoPDF();
					frame.setVisible(true);

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public GerarLaudoPDF() {
		textRef = null;
		finalizado = false;
		linhasSelecionadas = Principal.table.getSelectedRows();
		setTitle("Gerar arquivo em PDF");
		setResizable(false);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 813, 603);
		// setAlwaysOnTop(true);
		/*
		 * // Definindo a posicao da janela Dimension screenSize =
		 * Toolkit.getDefaultToolkit().getScreenSize();
		 * 
		 * int tolerancia = -20; int x = (int) ((screenSize.getWidth() - getWidth()));
		 * int y = (int) (((screenSize.getHeight() - getHeight()) / 3) - tolerancia);
		 * setLocation(x, y);
		 */
		setLocationRelativeTo(null);
		setVisible(true);

		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel lblHelp = new JLabel("Atalhos");
		lblHelp.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblHelp.setHorizontalTextPosition(SwingConstants.LEFT);
		lblHelp.setHorizontalAlignment(SwingConstants.RIGHT);
		lblHelp.setIcon(new ImageIcon(img_help));
		lblHelp.setBounds(688, 0, 99, 23);
		lblHelp.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		lblHelp.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

			}

			@Override
			public void mouseEntered(MouseEvent e) {
				convolvedImage = new BufferedImage(20, 20, BufferedImage.TYPE_INT_RGB);

				setBrightnessFactor(1.5f);

				lblHelp.setIcon(new ImageIcon(convolvedImage));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				// Define a imagem escalada no JLabel
				lblHelp.setIcon(new ImageIcon(img_help));
			}

			@Override
			public void mouseReleased(MouseEvent e) {

			}
		});
		contentPane.add(lblHelp);

		JPanel panel = new JPanel();
		panel.setLayout(null);

		// Crie uma instância personalizada de TitledBorder com fonte normal
		TitledBorder titledBorder1 = new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Visualiza\u00E7\u00E3o - Preview : :", TitledBorder.CENTER, TitledBorder.TOP,
				new Font("Dialog", Font.PLAIN, 12) // Defina a fonte com estilo normal
		);
		panel.setBorder(titledBorder1);

		panel.setBounds(457, 19, 330, 507);
		contentPane.add(panel);

		lblNewLabel = new JLabel("");
		lblNewLabel.setBounds(10, 19, 310, 477);
		panel.add(lblNewLabel);

		JPanel panel_1 = new JPanel();
		panel_1.setLayout(null);
		// Crie uma instância personalizada de TitledBorder com fonte normal
		TitledBorder titledBorder2 = new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Prepara\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP,
				new Font("Dialog", Font.PLAIN, 12) // Defina a fonte com estilo normal
		);
		panel_1.setBorder(titledBorder2);
		panel_1.setBounds(10, 19, 437, 507);
		contentPane.add(panel_1);

		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico *");
		lblNomeDoTcnico.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblNomeDoTcnico.setBounds(10, 26, 131, 29);
		panel_1.add(lblNomeDoTcnico);

		txtNomeTecnico = new JTextField();
		txtNomeTecnico.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtNomeTecnico.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		txtNomeTecnico.setColumns(10);
		txtNomeTecnico.setBounds(10, 48, 415, 29);
		panel_1.add(txtNomeTecnico);

		txtUsuarioRede = new JTextField();
		txtUsuarioRede.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtUsuarioRede.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		txtUsuarioRede.setColumns(10);
		txtUsuarioRede.setBounds(10, 96, 113, 29);
		panel_1.add(txtUsuarioRede);

		JLabel lblUsurioDeRede = new JLabel("Usuário de rede *");
		lblUsurioDeRede.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblUsurioDeRede.setBounds(10, 77, 113, 21);
		panel_1.add(lblUsurioDeRede);

		txtCentroCusto = new JTextField();
		txtCentroCusto.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtCentroCusto.setEnabled(false);
		txtCentroCusto.setText("321 - Sistemas CAT");
		txtCentroCusto.setColumns(10);
		txtCentroCusto.setBounds(144, 96, 281, 29);
		panel_1.add(txtCentroCusto);

		JLabel lblCentroDeCusto = new JLabel("Centro de custo *");
		lblCentroDeCusto.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblCentroDeCusto.setBounds(144, 77, 139, 21);
		panel_1.add(lblCentroDeCusto);

		JLabel lblTcnicoResponsvel = new JLabel("Técnico responsável");
		lblTcnicoResponsvel.setFont(new Font("Dialog", Font.BOLD, 13));
		lblTcnicoResponsvel.setBounds(10, 11, 131, 21);
		panel_1.add(lblTcnicoResponsvel);

		JSeparator separator = new JSeparator();
		separator.setBounds(9, 132, 416, 2);
		panel_1.add(separator);

		JLabel lblAnalise = new JLabel("Análise *");
		lblAnalise.setFont(new Font("Dialog", Font.BOLD, 12));
		lblAnalise.setBounds(10, 144, 93, 29);
		panel_1.add(lblAnalise);

		JSeparator separator_1 = new JSeparator();
		separator_1.setBounds(10, 305, 415, 2);
		panel_1.add(separator_1);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);

		scrollPane.setBounds(10, 176, 415, 108);
		panel_1.add(scrollPane);

		textAreaAnalise = new JTextArea();
		textAreaAnalise.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		scrollPane.setViewportView(textAreaAnalise);
		textAreaAnalise.setLineWrap(true);// Faz com que o texto quebre para a próxima linha
		textAreaAnalise.setWrapStyleWord(true);
		textAreaAnalise.setAutoscrolls(false);

		doc = new DefaultStyledDocument();
		doc.setDocumentFilter(new TextAreaSizeFilter(textAreaCaracteresLimit)); // Define o limite de qtd. caracteres
		;// Limite de 315 caracteres no JTextArea
		doc.addDocumentListener(new DocumentListener() {
			@Override
			public void changedUpdate(DocumentEvent e) {
				updateCount();
			}

			@Override
			public void insertUpdate(DocumentEvent e) {
				updateCount();
			}

			@Override
			public void removeUpdate(DocumentEvent e) {
				updateCount();
			}
		});
		SimpleAttributeSet attribs = new SimpleAttributeSet();
		StyleConstants.setAlignment(attribs, StyleConstants.ALIGN_JUSTIFIED);
		doc.setParagraphAttributes(0, doc.getLength(), attribs, false);
		textAreaAnalise.setDocument(doc);
		updateCount();

		comboBoxTemplate = new JComboBox();
		comboBoxTemplate.setFont(new Font("Dialog", Font.PLAIN, 12));
		comboBoxTemplate.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		comboBoxTemplate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (comboBoxTemplate.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		comboBoxTemplate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// Computador Lento (Destkop/notebook)
				if (comboBoxTemplate.getSelectedIndex() == 1) {
					textAreaAnalise.setText("");
					textAreaAnalise.setText(
							"Foi realizado uma análise na máquina do colaborador(a), cujo o mesmo se apresentou"
									+ " muito lento nos testes realizados, portanto, a máquina atual não atende os requisitos minímos"
									+ " para realizar atividades na empresa HPE prejudicando no desempenho profissional do colaborador(a).");
					textAreaAnalise.setCaretPosition(0);// Sobe para cima a barra de rolagem vertical\
					updateCount();
				}
				// Fonte (Desktop)
				else if (comboBoxTemplate.getSelectedIndex() == 2) {
					textAreaAnalise.setText("");
					textAreaAnalise
							.setText("Foi realizado uma análise na máquina do colaborador(a), cujo o mesmo apresentou"
									+ " problemas significativos na fonte, impossibilitando na inicialização da máquina, portanto,"
									+ " este componente é indispensável para o funcionamento da máquina para que o colaborador(a) possa"
									+ " estar realizando suas atividades na empresa HPE.");
					textAreaAnalise.setCaretPosition(0);// Sobe para cima a barra de rolagem vertical\
					updateCount();
				}
				// Bateria não segura carga (Notebook)
				else if (comboBoxTemplate.getSelectedIndex() == 3) {
					textAreaAnalise.setText("");
					textAreaAnalise
							.setText("Foi realizado uma análise no notebook do colaborador(a), cujo o mesmo apresentou"
									+ " severos problemas com a bateria, onde a mesma não segura carga suficiente para uso autônomo, portanto,"
									+ " este componente é indispensável para o funcionamento do notebook à longo alcance de tomadas (fonte) para que o colaborador(a) possa"
									+ " estar locomovendo-se e realizando suas atividades na empresa HPE sem precisar estar com o notebook carregando na tomada.");
					textAreaAnalise.setCaretPosition(0);// Sobe para cima a barra de rolagem vertical\
					updateCount();
				} else {
					comboBoxTemplate.setForeground(Color.LIGHT_GRAY);
					textAreaAnalise.setText("");
				}
			}
		});
		comboBoxTemplate.setModel(new DefaultComboBoxModel(new String[] { "Selecione uma template (Opcional)",
				"Computador lento", "Fonte", "Bateria não segura carga" }));
		comboBoxTemplate.setMaximumRowCount(7);
		comboBoxTemplate.setForeground(Color.LIGHT_GRAY);
		comboBoxTemplate.setBounds(154, 141, 271, 24);
		panel_1.add(comboBoxTemplate);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(11, 337, 414, 124);
		panel_1.add(scrollPane_1);

		editorPaneConsideracoesTecnicas = new JEditorPane();
		editorPaneConsideracoesTecnicas.setFont(new Font("Dialog", Font.PLAIN, 11));
		editorPaneConsideracoesTecnicas.setEnabled(false);
		editorPaneConsideracoesTecnicas.setText(
				"Considerando as análises realizadas, sugerimos a aquisição dos itens abaixo para upgrade/melhoria do computador:");
		consideracoesTecnicasTEMP = editorPaneConsideracoesTecnicas.getText();

		for (int i = 0; i < linhasSelecionadas.length; i++) {
			String textoAntigo = editorPaneConsideracoesTecnicas.getText();
			if (i == linhasSelecionadas.length - 1)
				editorPaneConsideracoesTecnicas
						.setText(textoAntigo + "\n\t•  " + "0" + Principal.qtd.get(linhasSelecionadas[i]) + " "
								+ Principal.item.get(linhasSelecionadas[i]) + ".");
			else
				editorPaneConsideracoesTecnicas
						.setText(textoAntigo + "\n\t•  " + "0" + Principal.qtd.get(linhasSelecionadas[i]) + " "
								+ Principal.item.get(linhasSelecionadas[i]) + ";");
			tempText = editorPaneConsideracoesTecnicas.getText();
		}
		scrollPane_1.setViewportView(editorPaneConsideracoesTecnicas);

		JLabel lblConsideracoesTecnicas = new JLabel("Considerações Técnicas *");
		lblConsideracoesTecnicas.setFont(new Font("Dialog", Font.BOLD, 12));
		lblConsideracoesTecnicas.setBounds(10, 312, 181, 29);
		panel_1.add(lblConsideracoesTecnicas);

		remaningLabel.setHorizontalAlignment(SwingConstants.RIGHT);
		remaningLabel.setFont(new Font("Dialog", Font.PLAIN, 12));
		remaningLabel.setBounds(217, 284, 208, 14);
		panel_1.add(remaningLabel);

		// "Remover -"
		btnDel = new JButton("-");
		btnDel.setFont(new Font("Dialog", Font.PLAIN, 12));
		btnDel.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		btnDel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				editorPaneConsideracoesTecnicas.setText(
						editorPaneConsideracoesTecnicas.getText().replaceFirst("• " + linkExibicao.get(num - 1), ""));

				if (editorPaneConsideracoesTecnicas.getText().endsWith("\n")) {
					for (int i = 0; i < linkExibicao.size(); i++)
						editorPaneConsideracoesTecnicas.getText().trim();
				}
				linkExibicao.remove(num - 1);
				linkEndereco.remove(num - 1);
				num--;
				if (linkExibicao.size() == 0) {
					btnDel.setEnabled(false);
					chckbxNewCheckBox.doClick();
					linkEndereco.clear();
					linkExibicao.clear();
				}
			}
		});
		btnDel.setBounds(203, 472, 47, 23);
		btnDel.setEnabled(false);
		panel_1.add(btnDel);

		// "Adicionar +"
		btnAdd = new JButton("+");
		btnAdd.setFont(new Font("Dialog", Font.PLAIN, 12));
		btnAdd.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		btnAdd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String temp = editorPaneConsideracoesTecnicas.getText().trim();

				// Criando a janela de multiplas opcoes
				JTextField txtExibicao = new JTextField();
				final JTextField txtEndereco = new JTextField();
				boolean flag = true;

				Object[] inputFields = { "Texto para exibição", txtExibicao, "Endereço (link)", txtEndereco };

				while (flag) {
					int option = JOptionPane.showConfirmDialog(null, inputFields, "Inserir Hyperlink",
							JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE);

					if (option == JOptionPane.OK_OPTION) {
						if (txtEndereco.getText().equals("") || txtExibicao.getText().equals("")) {
							JOptionPane.showMessageDialog(null, "Existem campos em vazio! Por favor preencha.");
						} else {
							linkExibicao.add(txtExibicao.getText());
							linkEndereco.add(txtEndereco.getText());
							flag = !flag;
						}
					} else {
						flag = !flag;
						if (num == 0) {
							btnAdd.setEnabled(false);
							chckbxNewCheckBox.doClick();
							linkEndereco.clear();
							linkExibicao.clear();
						}
					}
				}

				editorPaneConsideracoesTecnicas.setText(temp + "\n\t• " + linkExibicao.get(num));
				num++;

				btnDel.setEnabled(true);
			}
		});
		btnAdd.setBounds(154, 472, 47, 23);
		btnAdd.setEnabled(false);
		panel_1.add(btnAdd);

		// Inserir links
		chckbxNewCheckBox = new JCheckBox("Links de referência");
		chckbxNewCheckBox.setFont(new Font("Dialog", Font.PLAIN, 12));
		chckbxNewCheckBox.setSelected(false);
		chckbxNewCheckBox.addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});
		chckbxNewCheckBox.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				textRef = "Link para referência do mesmo:";
				if (chckbxNewCheckBox.isSelected()) {
					btnAdd.setEnabled(true);
					btnDel.setEnabled(false);
					tempText = editorPaneConsideracoesTecnicas.getText();
					editorPaneConsideracoesTecnicas.setText(tempText + "\n\n" + textRef);
					flagCheckBox = true;
					btnAdd.doClick();
				} else {
					textRef = null;
					btnAdd.setEnabled(false);
					btnDel.setEnabled(false);
					editorPaneConsideracoesTecnicas.setText(tempText);
					flagCheckBox = false;
				}
			}
		});
		chckbxNewCheckBox.setBounds(10, 472, 131, 23);
		panel_1.add(chckbxNewCheckBox);

		JLabel lblNewLabel_1 = new JLabel("Beta! Tem que testar");
		lblNewLabel_1.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblNewLabel_1.setForeground(Color.GRAY);
		lblNewLabel_1.setBounds(267, 476, 158, 14);
		panel_1.add(lblNewLabel_1);

		btnGerarArquivoPDF = new JButton("Gerar arquivo em PDF");
		btnGerarArquivoPDF.setFont(new Font("Dialog", Font.PLAIN, 12));
		btnGerarArquivoPDF.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnGerarArquivoPDF.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (textRef != null && num == 0) {
					JOptionPane.showMessageDialog(null,
							"Você marcou a caixa de seleção de inserir hyperlinks mas não inseriu nenhum link. Se não for usar, por favor desmarque.");
				} else if (txtNomeTecnico.getText().equals("") || txtUsuarioRede.getText().equals("")
						|| txtCentroCusto.getText().equals("") || textAreaAnalise.getText().equals("")) {
					JOptionPane.showMessageDialog(null,
							"Existem campos obrigatórios em brancos ou vazios, por favor preencha-os.");
				} else if (textAreaAnalise.getText().length() <= 30) {
					JOptionPane.showMessageDialog(null,
							"O campo da analise não pode passar de 440 caracteres e nem ser inferior a 30 caracteres.");
				} else {
					processamentoWord();
				}
			}
		});
		btnGerarArquivoPDF.setBounds(83, 531, 245, 23);
		contentPane.add(btnGerarArquivoPDF);

		btnVisualizar = new JButton("Visualizar");
		btnVisualizar.setFont(new Font("Dialog", Font.PLAIN, 12));
		btnVisualizar.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnVisualizar.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnVisualizar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					Desktop.getDesktop().open(new File(pdfPATH));
				} catch (IOException ex) {
					JOptionPane.showMessageDialog(null, ex);
				}
			}
		});
		btnVisualizar.setEnabled(false);
		btnVisualizar.setBounds(486, 531, 123, 23);
		contentPane.add(btnVisualizar);

		btnAbrirLocal = new JButton("Abrir local");
		btnAbrirLocal.setFont(new Font("Dialog", Font.PLAIN, 12));
		btnAbrirLocal.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnAbrirLocal.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnAbrirLocal.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					Desktop.getDesktop().open(new File(pdfPATH).getParentFile());

				} catch (IOException e1) {
					JOptionPane.showMessageDialog(null, e1);
				}
			}
		});
		btnAbrirLocal.setEnabled(false);
		btnAbrirLocal.setBounds(629, 530, 123, 23);
		contentPane.add(btnAbrirLocal);

		addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnGerarArquivoPDF.doClick();
				}
			}
		});

		// Adiciona a funcionalidade de desfazer/refazer em todos os componentes
		addUndoRedoFunctionality(getContentPane());

		txtNomeTecnico.setText(Principal.tecnico.get(Principal.table.getSelectedRow()));
	}

	private void processamentoWord() {
		try {
			// Shortcuts
			String fileName = "modelo laudo.docx";
			String userHome = System.getProperty("user.home");
			String pathRestante = "/ConversorXLSX-PDF/data/";

			// Preparacao do arquivo
			File file = new File(userHome + pathRestante + fileName); // Local pra preparar o arquivo
			FileInputStream fis = new FileInputStream(file.getAbsolutePath()); // Local pra preparar o arquivo
			XWPFDocument document = new XWPFDocument(fis); // Prepara o arquivo

			// Preparando campos de textos
			// Pegando dados de outras classes
			String laudo = Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]);
			String analise = GerarLaudoPDF.textAreaAnalise.getText();
			String consideracoes = consideracoesTecnicasTEMP;

			// Subistituindo as palavras do bookmarks...
			substituirFormularioDoCampoDeTexto(document, "Texto1", laudo);
			substituirFormularioDoCampoDeTexto(document, "Texto2", analise);
			substituirFormularioDoCampoDeTexto(document, "Texto3", consideracoes);

			// Instacias pra leitura da tabela do documento word
			XWPFTable table;
			XWPFTableRow row;
			XWPFTableCell cell;

			// --------- TABELA 1 - TECNICO RESPONSAVEL
			table = document.getTables().get(0); // Obtém a PRIMEIRA tabela no documento -> Padrão
			// Primeira linha da tabela
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha -> Padrao
			XWPFRun run = cell.getParagraphArray(0).createRun(); // Obtem a Run da cell
			run.setFontFamily("Calibri (Corpo)"); // Altera a fonte
			// Obtém o objeto CTRPr para a run usando o método getCTR() e isSetRPr(), ou
			// adiciona um novo objeto CTRPr caso ele não exista. Armazena a referência no
			// objeto rpr.
			CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			// Obtém o objeto CTHpsMeasure para o objeto CTRPr usando o método isSetSz() e
			// getSz(), ou adiciona um novo objeto CTHpsMeasure caso ele não exista.
			// Armazena a referência no objeto sz.
			CTHpsMeasure sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			// Altera o valor do tamanho da fonte para 10.5 * 2 = 21, usando o método
			// setVal() e passando o resultado como argumento.
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(txtNomeTecnico.getText()); // Nome completo

			// Segunda linha da tabela
			row = table.getRow(1);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(txtUsuarioRede.getText()); // Usuário de rede

			// Terceira linha da tabela
			row = table.getRow(2);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(txtCentroCusto.getText()); // Centro de custo

			// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

			// --------- TABELA 2 - USUARIO
			table = document.getTables().get(1); // Obtém a SEGUNDA tabela no documento -> Padrão
			// Primeira linha da tabela
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha -> Padrao
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.nomeSolicitante.get(linhasSelecionadas[0])); // Nome completo

			// Segunda linha da tabela
			row = table.getRow(1);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.usuario.get(linhasSelecionadas[0])); // Usuário de rede

			// Terceira linha da tabela
			row = table.getRow(2);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.centroCusto.get(linhasSelecionadas[0])); // Centro de custo

			// +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

			// --------- TABELA 3 - EQUIPAMENTO
			table = document.getTables().get(2); // Obtém a TERCEIRA tabela no documento -> Padrão
			// Primeira linha da tabela
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.dispositivo.get(linhasSelecionadas[0])); // Dispositivo
			cell = row.getCell(3); // Obtém a quarta célula da linha
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.hostname.get(linhasSelecionadas[0])); // Hostname

			// Segunda linha da tabela
			row = table.getRow(1);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.fabricante.get(linhasSelecionadas[0])); // Fabricante
			cell = row.getCell(3);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.modelo.get(linhasSelecionadas[0])); // Modelo

			// Terceira linha da tabela
			row = table.getRow(2);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.serviceTag.get(linhasSelecionadas[0])); // Service TAG
			cell = row.getCell(3);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.ativo.get(linhasSelecionadas[0])); // ID Ativo

			// Quarta linha da tabela
			row = table.getRow(3);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.dataAquisicao.get(linhasSelecionadas[0])); // Data de Aquisição
			cell = row.getCell(3);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10 * 2)));
			run.setText(Principal.cpu.get(linhasSelecionadas[0])); // CPU

			// Quinta linha da tabela
			row = table.getRow(4);
			cell = row.getCell(1);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.storage.get(linhasSelecionadas[0])); // Storage
			cell = row.getCell(3);
			run = cell.getParagraphArray(0).createRun();
			rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
			sz = rpr.isSetSz() ? rpr.getSz() : rpr.addNewSz();
			sz.setVal(BigInteger.valueOf((long) (10.5 * 2)));
			run.setText(Principal.memoria.get(linhasSelecionadas[0])); // Memoria

			// Criando hyperlink
			if (textRef != null && num != 0) {
				XWPFParagraph paragraph = document.createParagraph();
				run = paragraph.createRun();
				run.setText(textRef);

				for (int i = 0; i < num; i++) {
					paragraph = document.createParagraph();
					run = paragraph.createRun();

					// Atribuindo um ponto no inicio do paragrafo (final da run)
					CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
					cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
					CTLvl cTLvl = cTAbstractNum.addNewLvl();
					cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
					cTLvl.addNewLvlText().setVal("•");

					XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
					XWPFNumbering numbering = document.createNumbering();
					BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
					BigInteger numID = numbering.addNum(abstractNumID);

					// Printando 1 por 1
					if (i == num - 1) {
						run = paragraph.createHyperlinkRun(linkEndereco.get(i));
						run.setText(linkExibicao.get(i));
						run.setUnderline(UnderlinePatterns.SINGLE);
						run.setColor("0000FF");
						run = paragraph.createRun();
						run.setText(".");
						// Configurando font e realizando setText
						run.setBold(true);
						run.setFontFamily("Calibri (Corpo)");
						run.setFontSize(11);
					} else {
						run = paragraph.createHyperlinkRun(linkEndereco.get(i));
						run.setText(linkExibicao.get(i));
						run.setUnderline(UnderlinePatterns.SINGLE);
						run.setColor("0000FF");
						run = paragraph.createRun();
						run.setText(";");
						// Configurando font e realizando setText
						run.setBold(true);
						run.setFontFamily("Calibri (Corpo)");
						run.setFontSize(11);
					}

					// Margens
					paragraph.setNumID(numID);
					paragraph.setIndentationFirstLine(720); // 720 twips é aproximadamente 1 cm
				}
			}

			// Separando o number do ativo com o local HPE ou BW&P
			String[] ativo = Principal.ativo.get(GerarLaudoPDF.linhasSelecionadas[0]).split(" ");

			// ------- SALVAMENTO E CONVERSAO --------
			// Salvando o BACKUP (Word)...
			// Criando a pasta backup
			File pathBackup = new File(userHome + pathRestante + "Backup");
			ConfigManager.setConfigLine(3, pathBackup.toPath().toString());
			pathBackup.mkdirs();
			try (FileOutputStream out = new FileOutputStream(pathBackup.getPath() + "\\"
					+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
					+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".docx")) {
				// Saida do arquivo Word
				document.write(out);
				out.close();
				document.close();
			}

			// Conversao do documento word para pdf..
			// File diretorioInicial = new File("\\\\fscatorg01\\..."); //Substituir o local
			// do path do pdf
			File pathPdfGerados = new File(userHome + pathRestante + "Pdf generated");
			pathPdfGerados.mkdirs();
			
			// Abrindo o arquivo word a ser convertido
			try(	// 1 - Carregando arquivo WORD (DOCX)
					InputStream docxInputStream = new FileInputStream(pathBackup.getPath() + "\\"
						+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
						+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".docx");
					// 2 - Definindo posicao saida para o arquivo PDF
					OutputStream outputStream = new FileOutputStream(pathPdfGerados.getPath() + "\\"
							+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
							+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".pdf");){
				
				long start = System.currentTimeMillis();
				setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
				
				
				// Convertendo docx para pdf
				String pathTemp = System.getProperty("java.io.tmpdir") + "/LocalConverter";
				File temp = new File(pathTemp);
				temp.mkdir();
			
				IConverter converter = LocalConverter.builder().baseFolder(new File(pathTemp))
						.workerPool(20, 150, 2, TimeUnit.SECONDS).processTimeout(5, TimeUnit.SECONDS).build();

				@SuppressWarnings("unused")
				Future<Boolean> conversion = converter.convert(docxInputStream).as(DocumentType.MS_WORD).to(outputStream).as(DocumentType.PDF)
						.schedule();
				
				converter.shutDown();

				pdfPATH = pathPdfGerados.getPath() + "\\" + Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0])
						+ " - " + ativo[0] + " - " + Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0])
						+ ".pdf";
				
				
				// 4 - Gerando a preview da primeira pagina PDF para a Label...
				File arquivo = new File(pdfPATH);
				try(PDDocument pddDoc =  Loader.loadPDF(arquivo) ){
		            PDFRenderer pr = new PDFRenderer (pddDoc );
		            BufferedImage bim = pr.renderImageWithDPI(0, 300, ImageType.RGB);
					ImageIcon img = new ImageIcon(bim);
					ImageIcon imageIcon = new ImageIcon(img.getImage().getScaledInstance(310, 477, Image.SCALE_SMOOTH));
					lblNewLabel.setIcon(imageIcon);
		        } catch (IOException  e) {
		            e.printStackTrace();
		        }

				// Habilita botoes
				btnVisualizar.setEnabled(true);
				btnAbrirLocal.setEnabled(true);

				JOptionPane.showMessageDialog(null,
						"A conversão do arquivo MS Word para PDF ocorreu com sucesso!\n"
								+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
								+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".pdf"
								+ "\nDuração: " + ((System.currentTimeMillis() - start) / 1000) + " segundos.");
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
		    } catch (FileNotFoundException e) {
		        JOptionPane.showMessageDialog(null, "Arquivo DOCX não encontrado: " + e.getMessage());
		    } catch (IOException e) {
		        JOptionPane.showMessageDialog(null, "Erro de IO ao manipular o arquivo DOCX ou PDF: " + e.getMessage());
		    } catch (Exception e) {
		        JOptionPane.showMessageDialog(null, "Erro inesperado: " + e.getMessage());
		    }

		} catch (FileNotFoundException e) {
		    JOptionPane.showMessageDialog(null, "Erro ao criar diretório de backup: " + e.getMessage());
		} catch (IOException e) {
		    JOptionPane.showMessageDialog(null, "Erro de IO ao criar diretório de backup: " + e.getMessage());
		} catch (Exception e) {
		    JOptionPane.showMessageDialog(null, "Erro inesperado ao criar diretório de backup: " + e.getMessage());
		}
	}

	private static void substituirFormularioDoCampoDeTexto(XWPFDocument document, String keyword, String text) {
		// Flags
		boolean achouCampoDeFormulario = false;
		boolean quebraLinha = false;
		boolean fazerBulletList = false;

		// Analisando o XML do arquivo WORD
		for (XWPFParagraph paragraph : document.getParagraphs()) {
			for (XWPFRun run : paragraph.getRuns()) {
				// JOptionPane.showMessageDialog(null, run.getParagraph().getText());
				XmlCursor cursor = run.getCTR().newCursor();
				cursor.selectPath(
						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/@w:fldCharType");
				while (cursor.hasNextSelection()) {
					cursor.toNextSelection();
					XmlObject obj = cursor.getObject();
					if ("begin".equals(((SimpleValue) obj).getStringValue())) {
						cursor.toParent();
						obj = cursor.getObject();
						obj = obj.selectPath(
								"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:name/@w:val")[0];
						if (keyword.equals(((SimpleValue) obj).getStringValue())) {
							achouCampoDeFormulario = true;
						} else {
							achouCampoDeFormulario = false;
						}
					} else if ("end".equals(((SimpleValue) obj).getStringValue())) {
						if (achouCampoDeFormulario)
							return;
						achouCampoDeFormulario = false;
					}
				}
				if (achouCampoDeFormulario && run.getCTR().getTList().size() > 0) {
					run.getCTR().getTList().get(0).setStringValue("");
					achouCampoDeFormulario = false;
					quebraLinha = true;
					break;
				}
			}

			if (quebraLinha && keyword.equals("Texto1")) {
				paragraph.createRun().setText(text);
				quebraLinha = false;
			}
			if (quebraLinha && keyword.equals("Texto2")) {
				paragraph.createRun().setText(text);
				quebraLinha = false;
			}
			if (quebraLinha && keyword.equals("Texto3")) {
				paragraph.createRun().setText(text);
				quebraLinha = false;
				fazerBulletList = true;
			}
		}

		if (fazerBulletList) {
			for (int i = 0; i < linhasSelecionadas.length; i++) {
				XWPFParagraph paragraph = document.createParagraph();
				XWPFRun run = paragraph.createRun();
				// run.addTab();

				// Atribuindo um ponto no inicio do paragrafo (final da run)
				CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
				cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
				CTLvl cTLvl = cTAbstractNum.addNewLvl();
				cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
				cTLvl.addNewLvlText().setVal("•");

				XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
				XWPFNumbering numbering = document.createNumbering();
				BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
				BigInteger numID = numbering.addNum(abstractNumID);

				// Configurando font e realizando setText
				run.setFontFamily("Calibri (Corpo)");
				run.setFontSize(11);

				// Printando 1 por 1
				if (i == linhasSelecionadas.length - 1) {
					run.setBold(false);
					run.setText("0" + Principal.qtd.get(linhasSelecionadas[i]) + " ");
					run = paragraph.createRun();
					run.setBold(true);
					run.setText(Principal.item.get(linhasSelecionadas[i]) + ".");
				} else {
					run.setBold(false);
					run.setText("0" + Principal.qtd.get(linhasSelecionadas[i]) + " ");
					run = paragraph.createRun();
					run.setBold(true);
					run.setText(Principal.item.get(linhasSelecionadas[i]) + ";");
				}
				// Margens
				paragraph.setNumID(numID);
				paragraph.setIndentationFirstLine(720); // 720 twips é aproximadamente 1 cm
			}

		}
	}

	private void updateCount() {
		remaningLabel.setText((textAreaCaracteresLimit - doc.getLength()) + " caracteres restantes");
	}

	public static void addUndoRedoFunctionality(Container container) {
		for (Component component : container.getComponents()) {
			if (component instanceof JTextField) {
				JTextField textField = (JTextField) component;
				UndoManager undoManager = new UndoManager();
				Document docTemp = textField.getDocument();
				docTemp.addUndoableEditListener(undoManager);

				textField.addKeyListener(new KeyAdapter() {
					public void keyPressed(KeyEvent e) {
						if (e.getKeyCode() == KeyEvent.VK_Z && e.isControlDown()) {
							if (undoManager.canUndo()) {
								undoManager.undo();
							}
						} else if (e.getKeyCode() == KeyEvent.VK_Y && e.isControlDown()) {
							if (undoManager.canRedo()) {
								undoManager.redo();
							}
						}
					}
				});
			} else if (component instanceof JTextArea) {
				JTextArea textArea = (JTextArea) component;
				UndoManager undoManager = new UndoManager();
				Document docTemp = textArea.getDocument();
				docTemp.addUndoableEditListener(undoManager);

				textArea.addKeyListener(new KeyAdapter() {

					public void keyPressed(KeyEvent e) {
						if (e.getKeyCode() == KeyEvent.VK_Z && e.isControlDown()) {
							if (undoManager.canUndo()) {
								undoManager.undo();
							}
							if (textAreaAnalise.getText().equals("")) {
								comboBoxTemplate.setSelectedIndex(0);
								undoManager.die();
							}
						} else if (e.getKeyCode() == KeyEvent.VK_Y && e.isControlDown()) {
							if (undoManager.canRedo()) {
								undoManager.redo();
							}
						}
					}

				});
			} else if (component instanceof Container) {
				addUndoRedoFunctionality((Container) component);
			}
		}
	}

	static XWPFHyperlinkRun createHyperlinkRun(XWPFParagraph paragraph, String uri) {
		String rId = paragraph.getDocument().getPackagePart()
				.addExternalRelationship(uri, XWPFRelation.HYPERLINK.getRelation()).getId();

		CTHyperlink cthyperLink = paragraph.getCTP().addNewHyperlink();
		cthyperLink.setId(rId);
		cthyperLink.addNewR();

		return new XWPFHyperlinkRun(cthyperLink, cthyperLink.getRArray(0), paragraph);
	}

	private class KeyHandler implements KeyListener {
		public void keyPressed(KeyEvent e) {
			// código para executar quando uma tecla é pressionada
		}

		public void keyReleased(KeyEvent e) {
			// código para executar quando uma tecla é liberada
		}

		public void keyTyped(KeyEvent e) {
			// código para executar quando uma tecla é digitada
		}
	}

	private void setBrightnessFactor(float multiple) {
		float[] brightKernel = { multiple };
		BufferedImageOp bright = new ConvolveOp(new Kernel(1, 1, brightKernel));
		BufferedImage bufferedImage = new BufferedImage(20, 20, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2d = bufferedImage.createGraphics();
		g2d.drawImage(img_help, 0, 0, null);
		g2d.dispose();
		convolvedImage = bright.filter(bufferedImage, null);
		repaint();

	}
}
