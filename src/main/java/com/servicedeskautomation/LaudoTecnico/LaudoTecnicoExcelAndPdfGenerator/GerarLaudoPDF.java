package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.formdev.flatlaf.FlatIntelliJLaf;

public class GerarLaudoPDF extends JFrame {
	private static final long serialVersionUID = 4893492449132639712L;
	private JPanel contentPane;
	private JTextField txtNomeTecnico;
	private JTextField txtUsuarioRede;
	private JTextField txtCentroCusto;

	private JEditorPane editorPaneConsideracoesTecnicas;
	private JEditorPane editorPaneAnalise;

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

		try {
			int[] linhasSelecionadas = Principal.table.getSelectedRows();
			String fileName = "modelo laudo.docx";
			File file = new File(System.getProperty("user.home") + "/Documents/ConversorXLSX-PDF/data/" + fileName);
			FileInputStream fis;
			fis = new FileInputStream(file.getAbsolutePath());
			try (XWPFDocument document = new XWPFDocument(fis)) {
				// Atalho code
				String shortcut1 = "##1##";
				String shortcut2 = "##2##";
				String shortcut3 = "##3##";
				int j = 0;
				
				// Nome completo
				XWPFTable table = document.getTables().get(0); // Obtém a primeira tabela no documento
				XWPFTableRow row = table.getRow(0); // Obtém a primeira linha da tabela

				XWPFTableCell cell = row.getCell(1); // Obtém a segunda célula da linha
				cell.setText(Principal.nomeSolicitante.get(linhasSelecionadas[0])); // Adiciona um novo valor

				for (int i = 0; i < linhasSelecionadas.length; i++) {
					String currentText = editorPaneConsideracoesTecnicas.getText();
					editorPaneConsideracoesTecnicas
							.setText(currentText + "    • 0" + Principal.qtd.get(linhasSelecionadas[i]) + " "
									+ Principal.item.get(linhasSelecionadas[i]) + "\n");
				}

				// Adicionando
				for (XWPFParagraph paragraph : document.getParagraphs()) {
					String text = paragraph.getText();
					JOptionPane.showMessageDialog(null, "Text -> " + text);
					if (text.contains(shortcut1)) {
						text = text.replace(shortcut1, Principal.laudo.get(j));
						paragraph.removeRun(0);
						paragraph.createRun().setText(text);
					}else if(text.contains(shortcut2)){
						text = text.replace(shortcut2, editorPaneAnalise.getText());
						paragraph.removeRun(0);
						paragraph.createRun().setText(text);
					}else if(text.contains(shortcut3)){
						text = text.replace(shortcut3, editorPaneConsideracoesTecnicas.getText());
						paragraph.removeRun(0);
						paragraph.createRun().setText(text);
					}
					j++;
				}

				fis.close();
				FileOutputStream fos = new FileOutputStream(file.getAbsolutePath());
				document.write(fos);
				fos.close();
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
