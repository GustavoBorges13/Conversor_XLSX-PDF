package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

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
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.formdev.flatlaf.FlatIntelliJLaf;

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
	private List<String> linkExibicao = new ArrayList<String>();
	private List<String> linkEndereco = new ArrayList<String>();

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(new FlatIntelliJLaf());
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

		finalizado = false;
		linhasSelecionadas = Principal.table.getSelectedRows();
		setTitle("Gerar arquivo em PDF");
		setResizable(false);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 770, 598);
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

		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Visualiza\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(414, 11, 330, 507);
		contentPane.add(panel);

		lblNewLabel = new JLabel("");
		lblNewLabel.setBorder(new LineBorder(new Color(0, 0, 0)));
		lblNewLabel.setBounds(10, 19, 310, 477);
		panel.add(lblNewLabel);

		JPanel panel_1 = new JPanel();
		panel_1.setLayout(null);
		panel_1.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Prepara\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_1.setBounds(10, 11, 399, 507);
		contentPane.add(panel_1);

		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico *");
		lblNomeDoTcnico.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoTcnico.setBounds(10, 26, 131, 29);
		panel_1.add(lblNomeDoTcnico);

		txtNomeTecnico = new JTextField();
		txtNomeTecnico.setForeground(Color.BLACK);
		txtNomeTecnico.setColumns(10);
		txtNomeTecnico.setBounds(10, 48, 379, 29);
		panel_1.add(txtNomeTecnico);

		txtUsuarioRede = new JTextField();
		txtUsuarioRede.setForeground(Color.BLACK);
		txtUsuarioRede.setColumns(10);
		txtUsuarioRede.setBounds(10, 96, 113, 29);
		panel_1.add(txtUsuarioRede);

		JLabel lblUsurioDeRede = new JLabel("Usuário de rede *");
		lblUsurioDeRede.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblUsurioDeRede.setBounds(10, 77, 93, 21);
		panel_1.add(lblUsurioDeRede);

		txtCentroCusto = new JTextField();
		txtCentroCusto.setEnabled(false);
		txtCentroCusto.setText("321 - Sistemas CAT");
		txtCentroCusto.setForeground(Color.BLACK);
		txtCentroCusto.setColumns(10);
		txtCentroCusto.setBounds(144, 96, 245, 29);
		panel_1.add(txtCentroCusto);

		JLabel lblCentroDeCusto = new JLabel("Centro de custo *");
		lblCentroDeCusto.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblCentroDeCusto.setBounds(144, 77, 93, 21);
		panel_1.add(lblCentroDeCusto);

		JLabel lblTcnicoResponsvel = new JLabel("Técnico responsável");
		lblTcnicoResponsvel.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblTcnicoResponsvel.setBounds(10, 11, 131, 29);
		panel_1.add(lblTcnicoResponsvel);

		JSeparator separator = new JSeparator();
		separator.setBounds(9, 132, 370, 2);
		panel_1.add(separator);

		JLabel lblAnalise = new JLabel("Análise *");
		lblAnalise.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblAnalise.setBounds(10, 144, 93, 29);
		panel_1.add(lblAnalise);

		JSeparator separator_1 = new JSeparator();
		separator_1.setBounds(10, 305, 379, 2);
		panel_1.add(separator_1);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);

		scrollPane.setBounds(10, 176, 379, 108);
		panel_1.add(scrollPane);

		textAreaAnalise = new JTextArea();
		scrollPane.setViewportView(textAreaAnalise);
		textAreaAnalise.setLineWrap(true);// Faz com que o texto quebre para a próxima linha
		textAreaAnalise.setWrapStyleWord(true);
		textAreaAnalise.setAutoscrolls(false);
		doc = new DefaultStyledDocument();
		doc.setDocumentFilter(new DocumentSizeFilter(textAreaCaracteresLimit)); // Define o limite de qtd. caracteres
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

		JComboBox comboBoxTemplate = new JComboBox();
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
					comboBoxTemplate.setForeground(Color.BLACK);
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
					comboBoxTemplate.setForeground(Color.BLACK);
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
					comboBoxTemplate.setForeground(Color.BLACK);
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
		comboBoxTemplate.setBounds(154, 141, 235, 24);
		panel_1.add(comboBoxTemplate);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(11, 337, 378, 124);
		panel_1.add(scrollPane_1);

		editorPaneConsideracoesTecnicas = new JEditorPane();
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
		lblConsideracoesTecnicas.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblConsideracoesTecnicas.setBounds(10, 312, 181, 29);
		panel_1.add(lblConsideracoesTecnicas);

		remaningLabel.setHorizontalAlignment(SwingConstants.RIGHT);
		remaningLabel.setFont(new Font("Tahoma", Font.PLAIN, 9));
		remaningLabel.setBounds(257, 284, 131, 14);
		panel_1.add(remaningLabel);

		btnDel = new JButton("-");
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
				}
			}
		});
		btnDel.setBounds(193, 472, 47, 23);
		btnDel.setEnabled(false);
		panel_1.add(btnDel);

		btnAdd = new JButton("+");
		btnAdd.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String temp = editorPaneConsideracoesTecnicas.getText().trim();

				// Criando a janela de multiplas opcoes
				JTextField txtExibicao = new JTextField();
				final JTextField txtEndereco = new JTextField();
				boolean flag = true;

				Object[] inputFields = { "Texto para exibição", txtExibicao, "Endereço", txtEndereco };

				while (flag) {
					int option = JOptionPane.showConfirmDialog(null, inputFields, "Inserir Hyperlink",
							JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE);

					if (option == JOptionPane.OK_OPTION) {
						if (txtEndereco.getText().equals("") || txtExibicao.getText().equals("")) {
							JOptionPane.showMessageDialog(null, "Existem campos em vazio! Por favor preencha.");
						} else {
							linkExibicao.add(txtExibicao.getText());
							linkEndereco.add(txtEndereco.getText());
							flag=!flag;
						}
					}else {
						flag=!flag;
					}
				}

				editorPaneConsideracoesTecnicas.setText(temp + "\n\t• " + linkExibicao.get(num));
				num++;

				btnDel.setEnabled(true);
			}
		});
		btnAdd.setBounds(144, 472, 47, 23);
		btnAdd.setEnabled(false);
		panel_1.add(btnAdd);

		JCheckBox chckbxNewCheckBox = new JCheckBox("Links de referência");
		chckbxNewCheckBox.addActionListener(new ActionListener() {
			String textRef = "Link para referência do mesmo:";

			public void actionPerformed(ActionEvent e) {
				if (chckbxNewCheckBox.isSelected()) {
					btnAdd.setEnabled(true);
					btnDel.setEnabled(false);
					tempText = editorPaneConsideracoesTecnicas.getText();
					editorPaneConsideracoesTecnicas.setText(tempText + "\n\n" + textRef);
					flagCheckBox = true;
				} else {
					btnAdd.setEnabled(false);
					btnDel.setEnabled(false);
					editorPaneConsideracoesTecnicas.setText(tempText);
					flagCheckBox = false;
				}
			}
		});
		chckbxNewCheckBox.setBounds(10, 472, 131, 23);
		panel_1.add(chckbxNewCheckBox);
		
		JLabel lblNewLabel_1 = new JLabel("Em construção!");
		lblNewLabel_1.setForeground(Color.GRAY);
		lblNewLabel_1.setBounds(257, 476, 132, 14);
		panel_1.add(lblNewLabel_1);

		// BOTAO GERAR ARQUIVO PDF
		JButton btnGerarArquivoPDF = new JButton("Gerar arquivo em PDF");
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
				if (txtNomeTecnico.getText().equals("") || txtUsuarioRede.getText().equals("")
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
		btnGerarArquivoPDF.setBounds(83, 523, 245, 23);
		contentPane.add(btnGerarArquivoPDF);

		btnVisualizar = new JButton("Visualizar");
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
		btnVisualizar.setBounds(443, 523, 123, 23);
		contentPane.add(btnVisualizar);

		btnAbrirLocal = new JButton("Abrir local");
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
		btnAbrirLocal.setBounds(586, 522, 123, 23);
		contentPane.add(btnAbrirLocal);

	}

	private void processamentoWord() {
		try {
			// Shortcuts
			String fileName = "modelo laudo.docx";
			String userHome = System.getProperty("user.home");
			String pathRestante = "/Documents/ConversorXLSX-PDF/data/";

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

			// Separando o number do ativo com o local HPE ou BW&P
			String[] ativo = Principal.ativo.get(GerarLaudoPDF.linhasSelecionadas[0]).split(" ");

			// ------- SALVAMENTO E CONVERSAO --------
			// Salvando o BACKUP (Word)...
			// Criando a pasta backup
			File pathBackup = new File(userHome + pathRestante + "backup");
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
			//File diretorioInicial = new File("\\\\fscatorg01\\..."); //Substituir o local do path do pdf
			File pathPdfGerados = new File(userHome + pathRestante + "Pdf generated");
			pathPdfGerados.mkdirs();

			// Abrindo o arquivo word a ser convertido
			try {
				long start = System.currentTimeMillis();
				setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

				// 1 - Carregando arquivo WORD (DOCX)
				InputStream docxInputStream = new FileInputStream(pathBackup.getPath() + "\\"
						+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
						+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".docx");
				// 2 - Definindo posicao saida para o arquivo PDF
				OutputStream outputStream = new FileOutputStream(pathPdfGerados.getPath() + "\\"
						+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
						+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".pdf");

				// 3 - Convertendo para PDF
				IConverter converter = LocalConverter.builder().build();
				converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF)
						.execute();
				outputStream.close();

				pdfPATH = pathPdfGerados.getPath() + "\\" + Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0])
						+ " - " + ativo[0] + " - " + Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0])
						+ ".pdf";
				File arquivo = new File(pdfPATH);

				// 4 - Gerando a preview da primeira pagina PDF para a Label...
				PDDocument doc = PDDocument.load(arquivo);
				org.apache.pdfbox.rendering.PDFRenderer pdfRenderer = new org.apache.pdfbox.rendering.PDFRenderer(doc);
				BufferedImage bim = pdfRenderer.renderImageWithDPI(0, 300, ImageType.RGB);
				ImageIcon img = new ImageIcon(bim);
				ImageIcon imageIcon = new ImageIcon(img.getImage().getScaledInstance(310, 477, Image.SCALE_SMOOTH));
				lblNewLabel.setIcon(imageIcon);

				// Habilita botoes
				btnVisualizar.setEnabled(true);
				btnAbrirLocal.setEnabled(true);

				JOptionPane.showMessageDialog(null,
						"A conversão do arquivo MS Word para PDF ocorreu com sucesso!\n"
								+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
								+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".pdf"
								+ "\nDuração: " + ((System.currentTimeMillis() - start) / 1000) + " segundos.");
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			} catch (Exception e) {
				JOptionPane.showMessageDialog(null, "Erro Exception IN: " + e);
			}

		} catch (

		FileNotFoundException e) {
			JOptionPane.showMessageDialog(null, "Erro com arquivo: " + e);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "Erro IO: " + e);
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "Erro Exception: " + e);
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
}
