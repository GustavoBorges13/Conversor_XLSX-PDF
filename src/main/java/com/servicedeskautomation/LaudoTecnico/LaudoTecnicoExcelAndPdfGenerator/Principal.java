package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableCellRenderer;
//XSSF = (XML SpreadSheet Format) – Used to reading and writting Open Office XML (XLSX) format files.   

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Principal extends JFrame {
	private static final long serialVersionUID = 6391163855934589017L;
	private JPanel contentPane;
	private JTextField jlocal;
	private JTextField jtitulo;
	private JTable table;
	@SuppressWarnings("unused")
	private DefaultTableCellRenderer renderer = new DefaultTableCellRenderer();
	private static int alinhamento = SwingConstants.LEFT;
	private File pathfile;
	private ArrayList<Integer> item = new ArrayList<Integer>();
	private ArrayList<String> desc = new ArrayList<String>();

	// Customize the code to set the background and foreground color for each column
	// of a JTable
	class ColumnColorRenderer extends DefaultTableCellRenderer {
		private static final long serialVersionUID = -2920708207060590184L;
		Color backgroundColor, foregroundColor;

		public ColumnColorRenderer(Color backgroundColor, Color foregroundColor) {
			super();
			this.backgroundColor = backgroundColor;
			this.foregroundColor = foregroundColor;
		}

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			Component cell = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			cell.setBackground(backgroundColor);
			cell.setForeground(foregroundColor);
			return cell;
		}
	}

	public class HorizontalAlignmentHeaderRenderer implements TableCellRenderer {
		private int horizontalAlignment = SwingConstants.LEFT;

		public HorizontalAlignmentHeaderRenderer(int horizontalAlignment) {
			this.horizontalAlignment = horizontalAlignment;
		}

		@Override
		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			TableCellRenderer r = table.getTableHeader().getDefaultRenderer();
			JLabel l = (JLabel) r.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			l.setHorizontalAlignment(horizontalAlignment);
			return l;
		}
	}

	@SuppressWarnings("static-access")
	public void preencherTabelaProprietario() {
		ArrayList<Object[]> dados = new ArrayList<>();
		String[] Colunas = new String[] { " Item", " Descricao" };

		for (int i = 0; i < item.size(); i++) {
			dados.add(new Object[] { (" " + item.get(i)), (" " + desc.get(i)) });
		}

		ModeloTabela modelo = new ModeloTabela(dados, Colunas);
		table.setModel(modelo);

		// Nao deixa a aumentar a largura das colunas da tabela usando o mouse e realiza
		// os alinhamentos das colunas e linhas!
		table.getColumnModel().getColumn(0).setPreferredWidth(160);
		table.getColumnModel().getColumn(0).setResizable(false);
		table.getColumnModel().getColumn(0).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(1).setPreferredWidth(540);
		table.getColumnModel().getColumn(1).setResizable(false);
		table.getColumnModel().getColumn(1).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));

		// Nao vai reodernar os nomes e titulos do cabeçalho da tabela
		table.getTableHeader().setReorderingAllowed(false);

		// Nao permite aumentar na tabela as colunas
		table.setAutoResizeMode(table.AUTO_RESIZE_OFF);

		// Faz com que o usuario selecione um dado na tabela POR VEZ
		table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

		// Realiza a configuracao de fontes
		table.getTableHeader().setFont(new Font("Dialog", Font.BOLD | Font.ITALIC, 12));
		table.setFont(new Font("Dialog", Font.PLAIN, 9));
	}

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Principal frame = new Principal();
					frame.setVisible(true);

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public Principal() {
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 704, 528);
		setLocationRelativeTo(null);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		jlocal = new JTextField();
		jlocal.setBounds(39, 49, 468, 39);
		contentPane.add(jlocal);
		jlocal.setColumns(10);

		JButton btnXLS = new JButton("XLS");

		// button choose file/escolher arquivo
		btnXLS.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fc = new JFileChooser();
				fc.setPreferredSize(new Dimension(700, 400));
				fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

				int resultado = fc.showOpenDialog(null);

				if (resultado == JFileChooser.CANCEL_OPTION) {
					// Para cancelar a janela se necessario
				} else {
					pathfile = fc.getSelectedFile();
					jlocal.setText(pathfile.toString().trim());
					jtitulo.setText(fc.getSelectedFile().getName());
				}
			}
		});
		btnXLS.setBounds(541, 49, 113, 39);
		contentPane.add(btnXLS);

		JButton btnPreencher = new JButton("Preencher");

		// Button fill/preencher
		btnPreencher.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				try {
					int i = 1;
					int pos = 0;
					// Abre o arquivo excel
					XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(pathfile));

					// Junta todas as planilhas deste arquivo.
					// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
					XSSFSheet sheet = work.getSheetAt(0);

					// Linha de referencia (selecionar a partir de uma determinada linha...)
					XSSFRow row = sheet.getRow(i);
					// XSSFRow row = null;

					// System.out.println("TESTE"); //Linha debug
					while ((row = sheet.getRow(i)) != null) {
						if (row.getCell(0).getNumericCellValue() != 0 && row.getCell(1).getStringCellValue() != null) {
							item.add((int) row.getCell(0).getNumericCellValue());
							desc.add(row.getCell(1).getStringCellValue());
							System.out.print("Laudo N :: " + item.get(pos));
							System.out.println("\tNome :: " + desc.get(pos));
							// System.out.println("TESTE");
							pos++;
							i++;
						} else {
							preencherTabelaProprietario();
							break;
						}

					}
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});

		btnPreencher.setBounds(541, 131, 113, 39);
		contentPane.add(btnPreencher);

		jtitulo = new JTextField();
		jtitulo.setColumns(10);
		jtitulo.setBounds(39, 131, 468, 39);
		contentPane.add(jtitulo);

		JLabel lblNewLabel = new JLabel("Titulo da planilha");
		lblNewLabel.setBounds(39, 106, 105, 14);
		contentPane.add(lblNewLabel);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(39, 195, 618, 248);
		contentPane.add(scrollPane);

		table = new JTable();
		scrollPane.setViewportView(table);

		preencherTabelaProprietario();

	}
}
