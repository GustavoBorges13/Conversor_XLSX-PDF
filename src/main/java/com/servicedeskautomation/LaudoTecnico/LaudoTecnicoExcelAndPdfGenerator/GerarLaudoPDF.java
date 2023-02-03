package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
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
import javax.swing.border.TitledBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;

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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import com.formdev.flatlaf.FlatIntelliJLaf;

public class GerarLaudoPDF extends JFrame {
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
		linhasSelecionadas = Principal.table.getSelectedRows();

		setResizable(false);
		setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
		setBounds(100, 100, 770, 632);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel lblTitulo = new JLabel("Gerar arquivo em PDF");
		lblTitulo.setHorizontalAlignment(SwingConstants.CENTER);
		lblTitulo.setFont(new Font("Tahoma", Font.PLAIN, 17));
		lblTitulo.setBounds(10, 11, 734, 21);
		contentPane.add(lblTitulo);

		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Visualiza\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(414, 43, 330, 506);
		contentPane.add(panel);

		JPanel panel_1 = new JPanel();
		panel_1.setLayout(null);
		panel_1.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Prepara\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_1.setBounds(10, 43, 399, 506);
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
									+ " severos problemas com a fonte, impossibilitando na inicialização da máquina, portanto,"
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
						.setText(textoAntigo + "\n\t•  " + Principal.item.get(linhasSelecionadas[i]) + ".");
			else
				editorPaneConsideracoesTecnicas
						.setText(textoAntigo + "\n\t•  " + Principal.item.get(linhasSelecionadas[i]) + ";");
		}
		scrollPane_1.setViewportView(editorPaneConsideracoesTecnicas);

		JLabel lblConsideracoesTecnicas = new JLabel("Considerações Técnicas *");
		lblConsideracoesTecnicas.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblConsideracoesTecnicas.setBounds(10, 312, 181, 29);
		panel_1.add(lblConsideracoesTecnicas);

		// BOTAO INSERIR LINK
		JButton btnInsirirLink = new JButton("Inserir link");
		btnInsirirLink.setBounds(10, 472, 89, 23);
		panel_1.add(btnInsirirLink);

		remaningLabel.setHorizontalAlignment(SwingConstants.RIGHT);
		remaningLabel.setFont(new Font("Tahoma", Font.PLAIN, 9));
		remaningLabel.setBounds(257, 284, 131, 14);
		panel_1.add(remaningLabel);

		// BOTAO GERAR ARQUIVO PDF
		JButton btnGerarArquivoPDF = new JButton("Gerar arquivo em PDF");
		btnGerarArquivoPDF.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				// Executa o metodo de substituicao de valores no documento word...
				// Abrindo o arquivo...
				processamentoWord();
			}
		});
		btnGerarArquivoPDF.setBounds(86, 560, 245, 23);
		contentPane.add(btnGerarArquivoPDF);

		// BOTAO VISUALIZAR
		JButton btnVisualizar = new JButton("Visualizar");
		btnVisualizar.setBounds(446, 560, 123, 23);
		contentPane.add(btnVisualizar);

		// BOTAO ABRIR LOCAL
		JButton btnAbrirLocal = new JButton("Abrir local");
		btnAbrirLocal.setBounds(589, 559, 123, 23);
		contentPane.add(btnAbrirLocal);

		// Definindo a posicao da janela
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();

		int tolerancia = -20;
		int x = (int) ((screenSize.getWidth() - getWidth()));
		int y = (int) (((screenSize.getHeight() - getHeight()) / 3) - tolerancia);
		setLocation(x, y);
		setVisible(true);
	}

	private static void processamentoWord() {
		try {
			// Shortcuts
			String fileName = "modelo laudo.docx";
			String userHome = System.getProperty("user.home");
			String pathRestante = "/Documents/ConversorXLSX-PDF/data/";

			// Preparacao do arquivo
			File file = new File(userHome + pathRestante + fileName); // Local pra preparar o arquivo
			FileInputStream fis = new FileInputStream(file.getAbsolutePath()); // Local pra preparar o arquivo
			XWPFDocument document = new XWPFDocument(fis); // Prepara o arquivo

			// Instacias pra leitura da tabela do documento word
			XWPFTable table;
			XWPFTableRow row;
			XWPFTableCell cell;

			
			// --------- TABELA 1 - TECNICO RESPONSAVEL
			table = document.getTables().get(0); // Obtém a PRIMEIRA tabela no documento -> Padrão
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha -> Padrao
			cell.setText(txtNomeTecnico.getText()); // Nome completo
			row = table.getRow(1);
			cell.setText(txtUsuarioRede.getText()); // Usuário de rede
			row = table.getRow(2);
			cell.setText(txtCentroCusto.getText()); // Centro de custo

			
			// --------- TABELA 2 - USUARIO
			table = document.getTables().get(1); // Obtém a SEGUNDA tabela no documento -> Padrão
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha -> Padrao
			cell.setText(Principal.nomeSolicitante.get(linhasSelecionadas[0])); // Nome completo
			row = table.getRow(1);
			cell.setText(Principal.usuario.get(linhasSelecionadas[0])); // Usuário de rede
			row = table.getRow(2);
			cell.setText(Principal.centroCusto.get(linhasSelecionadas[0])); // Centro de custo

			
			// --------- TABELA 3 - EQUIPAMENTO
			table = document.getTables().get(2); // Obtém a TERCEIRA tabela no documento -> Padrão
			// Primeira linha da tabela
			row = table.getRow(0); // Obtém a primeira linha da tabela
			cell = row.getCell(1); // Obtém a segunda célula da linha
			cell.setText(Principal.dispositivo.get(linhasSelecionadas[0])); // Dispositivo
			cell = row.getCell(3); // Obtém a quarta célula da linha
			cell.setText(Principal.hostname.get(linhasSelecionadas[0])); // Hostname
			// Segunda linha da tabela
			row = table.getRow(1);
			cell = row.getCell(1);
			cell.setText(Principal.fabricante.get(linhasSelecionadas[0])); // Fabricante
			cell = row.getCell(3);
			cell.setText(Principal.modelo.get(linhasSelecionadas[0])); // Modelo
			// Terceira linha da tabela
			row = table.getRow(2);
			cell = row.getCell(1);
			cell.setText(Principal.serviceTag.get(linhasSelecionadas[0])); // Service TAG
			cell = row.getCell(3);
			cell.setText(Principal.ativo.get(linhasSelecionadas[0])); // ID Ativo
			// Quarta linha da tabela
			row = table.getRow(3);
			cell = row.getCell(1);
			cell.setText(Principal.dataAquisicao.get(linhasSelecionadas[0])); // Data de Aquisição
			cell = row.getCell(3);
			cell.setText(Principal.cpu.get(linhasSelecionadas[0])); // CPU
			// Quinta linha da tabela
			row = table.getRow(4);
			cell = row.getCell(1);
			cell.setText(Principal.storage.get(linhasSelecionadas[0])); // Storage
			cell = row.getCell(3);
			cell.setText(Principal.memoria.get(linhasSelecionadas[0])); // Memoria

			
			// Preparando campos de textos
			// Pegando dados de outras classes
			String laudo = Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]);
			String analise = GerarLaudoPDF.textAreaAnalise.getText();
			String consideracoes = consideracoesTecnicasTEMP;
			JOptionPane.showMessageDialog(null, analise);

			// Subistituindo as palavras do bookmarks...
			substituirFormularioDoCampoDeTexto(document, "Texto1", laudo);
			substituirFormularioDoCampoDeTexto(document, "Texto2", analise);
			substituirFormularioDoCampoDeTexto(document, "Texto3", consideracoes);

			// Separando o number do ativo com o local HPE ou BW&P
			String[] ativo = Principal.ativo.get(GerarLaudoPDF.linhasSelecionadas[0]).split(" ");			
			
			// Local onde será baixado
			File folder = new File(userHome + pathRestante + "backup");
			folder.mkdirs();
			try (FileOutputStream out = new FileOutputStream(folder.getPath() + "\\"
					+ Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0] + " - "
					+ Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".docx")) {
				document.write(out);
				out.close();
				document.close();
			}
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(null, "Erro com arquivo: " + e);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "Erro: " + e);
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
					// run.getCTR().getTList().get(0).setStringValue(text);
					achouCampoDeFormulario = false;
					quebraLinha = true;
					break;
					// System.out.println(run.getCTR());
				}
			}

			if (quebraLinha && keyword.equals("Texto2")) {
				paragraph.createRun().setText("");
				quebraLinha = false;
				fazerBulletList = true;
			}
			if (quebraLinha && keyword.equals("Texto3")) {
				paragraph.createRun().setText("");
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
				if (i == linhasSelecionadas.length - 1)
					run.setText(Principal.item.get(linhasSelecionadas[i]) + ".");
				else
					run.setText(Principal.item.get(linhasSelecionadas[i]) + ";");

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
