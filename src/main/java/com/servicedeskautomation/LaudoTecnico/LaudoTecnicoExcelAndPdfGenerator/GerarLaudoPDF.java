package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;

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
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;

import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLevelSuffix;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat;

import com.formdev.flatlaf.FlatIntelliJLaf;

public class GerarLaudoPDF extends JFrame {
	private static final long serialVersionUID = 4893492449132639712L;
	private JPanel contentPane;
	private JTextField txtNomeTecnico;
	private JTextField txtUsuarioRede;
	private JTextField txtCentroCusto;
	static JEditorPane editorPaneConsideracoesTecnicas;
	static JEditorPane editorPaneAnalise;
	static int[] linhasSelecionadas;

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

	@SuppressWarnings({ "unused", "unchecked", "rawtypes" })
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
		panel.setBounds(414, 43, 330, 539);
		contentPane.add(panel);

		JPanel panel_1 = new JPanel();
		panel_1.setLayout(null);
		panel_1.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Prepara\u00E7\u00E3o : :", TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_1.setBounds(20, 43, 389, 539);
		contentPane.add(panel_1);

		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico *");
		lblNomeDoTcnico.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoTcnico.setBounds(10, 26, 131, 29);
		panel_1.add(lblNomeDoTcnico);

		txtNomeTecnico = new JTextField();
		txtNomeTecnico.setForeground(Color.BLACK);
		txtNomeTecnico.setColumns(10);
		txtNomeTecnico.setBounds(10, 48, 369, 29);
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
		txtCentroCusto.setText("321 - Sistemas CAT");
		txtCentroCusto.setForeground(Color.BLACK);
		txtCentroCusto.setColumns(10);
		txtCentroCusto.setBounds(144, 96, 235, 29);
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
		separator_1.setBounds(10, 305, 369, 2);
		panel_1.add(separator_1);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane.setBounds(10, 168, 369, 126);
		panel_1.add(scrollPane);

		editorPaneAnalise = new JEditorPane();
		scrollPane.setViewportView(editorPaneAnalise);

		JComboBox comboBoxTemplate = new JComboBox();
		comboBoxTemplate.setModel(new DefaultComboBoxModel(
				new String[] { "Selecione uma template (Opcional)", "Computador lento", "Fonte" }));
		comboBoxTemplate.setMaximumRowCount(7);
		comboBoxTemplate.setForeground(Color.LIGHT_GRAY);
		comboBoxTemplate.setBounds(154, 141, 224, 24);
		panel_1.add(comboBoxTemplate);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(11, 337, 367, 124);
		panel_1.add(scrollPane_1);

		editorPaneConsideracoesTecnicas = new JEditorPane();
		editorPaneConsideracoesTecnicas.setText(
				"Considerando as análises realizadas, sugerimos a aquisição dos itens abaixo para upgrade do computador:\r\n");
		scrollPane_1.setViewportView(editorPaneConsideracoesTecnicas);

		JLabel lblConsideracoesTecnicas = new JLabel("Considerações Técnicas *");
		lblConsideracoesTecnicas.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblConsideracoesTecnicas.setBounds(10, 312, 181, 29);
		panel_1.add(lblConsideracoesTecnicas);

		JButton btnInsirirLink = new JButton("Inserir link");
		btnInsirirLink.setBounds(10, 472, 89, 23);
		panel_1.add(btnInsirirLink);
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		// rect.getMaxX() - getWidth();

		int tolerancia = -20;
		int x = (int) ((screenSize.getWidth() - getWidth()));
		int y = (int) (((screenSize.getHeight() - getHeight()) / 3) - tolerancia);
		setLocation(x, y);
		setVisible(true);

		// Executa o metodo de substituicao de valores no documento word...
		// Abrindo o arquivo...
		try {
			String fileName = "modelo laudo.docx";
			String userHome = System.getProperty("user.home");
			String pathRestante = "/Documents/ConversorXLSX-PDF/data/";
			File file = new File(userHome + pathRestante + fileName);
			FileInputStream fis;
			fis = new FileInputStream(file.getAbsolutePath());
			XWPFDocument document = new XWPFDocument(fis);

			// Transcrevendo os itens para o campo de consideracoes tecnicas e mesclando...
			/*
			 * for (int i = 0; i < linhasSelecionadas.length; i++) { String currentText =
			 * editorPaneConsideracoesTecnicas.getText(); editorPaneConsideracoesTecnicas
			 * .setText(currentText + "    • 0" + Principal.qtd.get(linhasSelecionadas[i]) +
			 * " " + Principal.item.get(linhasSelecionadas[i]) + "\n"); }
			 */

			// Pegando dados de outras classes
			String laudo = Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]);
			String analise = GerarLaudoPDF.editorPaneAnalise.getText();
			String consideracoes = GerarLaudoPDF.editorPaneConsideracoesTecnicas.getText();

			// Subistituindo as palavras do bookmarks...
			replaceFormFieldText(document, "Texto1", laudo);
			replaceFormFieldText(document, "Texto2", analise);
			replaceFormFieldText(document, "Texto3", consideracoes);

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
			JOptionPane.showMessageDialog(null, "Erro: " + e);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "Erro: " + e);
		}
	}

	private static void replaceFormFieldText(XWPFDocument document, String keyword, String text) {
		boolean foundformfield = false;
		boolean quebraLinha = false;

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
							foundformfield = true;
						} else {
							foundformfield = false;
						}
					} else if ("end".equals(((SimpleValue) obj).getStringValue())) {
						if (foundformfield)
							return;
						foundformfield = false;
					}
				}
				if (foundformfield && run.getCTR().getTList().size() > 0) {
					run.getCTR().getTList().get(0).setStringValue("");

					// run.getCTR().getTList().get(0).setStringValue(text);
					foundformfield = false;
					quebraLinha = true;
					break;
					// System.out.println(run.getCTR());
				}
			}
			if (quebraLinha && keyword.equals("Texto3")) {
				// paragraph.removeRun(0);

				paragraph.insertNewRun(0).setText(text);
				for (int i = 0; i < linhasSelecionadas.length; i++) {
					XWPFParagraph paragraphNew = document.createParagraph();
					XWPFRun run = paragraphNew.createRun();

					// Bullet list
					CTAbstractNum cTAbstractNum = CTAbstractNum.Factory.newInstance();
					cTAbstractNum.setAbstractNumId(BigInteger.valueOf(0));
					CTLvl cTLvl = cTAbstractNum.addNewLvl();
					cTLvl.addNewNumFmt().setVal(STNumberFormat.BULLET);
					cTLvl.addNewLvlText().setVal("•");

					XWPFAbstractNum abstractNum = new XWPFAbstractNum(cTAbstractNum);
					XWPFNumbering numbering = document.createNumbering();
					BigInteger abstractNumID = numbering.addAbstractNum(abstractNum);
					BigInteger numID = numbering.addNum(abstractNumID);

					paragraphNew.setNumID(numID);
					// font size for bullet point in half pt
					//paragraph.getCTP().getPPr().addNewRPr().addNewSz().setVal(BigInteger.valueOf(48));
					run.setText(Principal.item.get(linhasSelecionadas[i]));
					run.setFontSize(11);
				}
			}
		}
	}
}
