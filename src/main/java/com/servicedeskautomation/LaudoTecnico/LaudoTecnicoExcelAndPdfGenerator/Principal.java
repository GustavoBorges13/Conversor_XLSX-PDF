package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
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
	private static int alinhamento = SwingConstants.LEFT;
	private File pathfile;

	private ArrayList<String> laudo = new ArrayList<String>();
	private ArrayList<String> nomeSolicitante = new ArrayList<String>();
	private ArrayList<String> usuario = new ArrayList<String>();
	private ArrayList<String> centroCusto = new ArrayList<String>();
	private ArrayList<String> item = new ArrayList<String>();
	private ArrayList<String> qtd = new ArrayList<String>();
	private ArrayList<String> ativo = new ArrayList<String>();
	private ArrayList<String> dispositivo = new ArrayList<String>();
	private ArrayList<String> hostname = new ArrayList<String>();
	private ArrayList<String> fabricante = new ArrayList<String>();
	private ArrayList<String> modelo = new ArrayList<String>();
	private ArrayList<String> serviceTag = new ArrayList<String>();
	private ArrayList<String> dataAquisicao = new ArrayList<String>();
	private ArrayList<String> cpu = new ArrayList<String>();
	private ArrayList<String> storage = new ArrayList<String>();
	private ArrayList<String> memoria = new ArrayList<String>();
	private ArrayList<String> tecnico = new ArrayList<String>();
	private ArrayList<String> observacao = new ArrayList<String>();
	private ArrayList<String> status = new ArrayList<String>();
	private ArrayList<String> colunas = new ArrayList<String>();

	/*
	 * Alinhador de fonts da tabela, deixa as fontes centralizadas para a lateral
	 * Para controlar mude o .LEFT para o que desejar... .CENTER, .RIGHT
	 */
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
		
		try {
			for (int i = 0; i < colunas.size(); i++) {
				dados.add(new Object[] { (" " + laudo.get(i)), (" " + nomeSolicitante.get(i)), (" " + usuario.get(i)),
						(" " + centroCusto.get(i)), (" " + item.get(i)), (" " + qtd.get(i)), (" " + ativo.get(i)),
						(" " + dispositivo.get(i)), (" " + hostname.get(i)), (" " + fabricante.get(i)),
						(" " + modelo.get(i)), (" " + serviceTag.get(i)), (" " + dataAquisicao.get(i)),
						(" " + cpu.get(i)), (" " + storage.get(i)), (" " + memoria.get(i)), (" " + tecnico.get(i)),
						(" " + observacao.get(i)), (" " + status.get(i)) });
			} 
		} catch (Exception e) {
			// TODO: handle exception
		}
		// ModeloTabela mode = new ModeloTabela(dados,Colunas);
		ModeloTabela modelo = new ModeloTabela(dados, colunas);
		table.setModel(modelo);
		
		// Nao deixa a aumentar a largura das colunas da tabela usando o mouse e realiza
		// os alinhamentos das colunas e linhas!
		table.getColumnModel().getColumn(0).setPreferredWidth(50); //coluna LAUDO
		table.getColumnModel().getColumn(0).setResizable(false);
		table.getColumnModel().getColumn(0).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(1).setPreferredWidth(170); //coluna NOME SOLICITANTE
		table.getColumnModel().getColumn(1).setResizable(false);
		table.getColumnModel().getColumn(1).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(2).setPreferredWidth(65); //coluna USUARIO
		table.getColumnModel().getColumn(2).setResizable(false);
		table.getColumnModel().getColumn(2).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(3).setPreferredWidth(200); //coluna CENTRO DE CUSTO
		table.getColumnModel().getColumn(3).setResizable(false);
		table.getColumnModel().getColumn(3).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(4).setPreferredWidth(190); //coluna ITEM
		table.getColumnModel().getColumn(4).setResizable(false);
		table.getColumnModel().getColumn(4).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(5).setPreferredWidth(30); //coluna QTD
		table.getColumnModel().getColumn(5).setResizable(false);
		table.getColumnModel().getColumn(5).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(6).setPreferredWidth(95); //coluna ATIVO
		table.getColumnModel().getColumn(6).setResizable(false);
		table.getColumnModel().getColumn(6).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(7).setPreferredWidth(80); //coluna DISPOSITIVO
		table.getColumnModel().getColumn(7).setResizable(false);
		table.getColumnModel().getColumn(7).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(8).setPreferredWidth(85); //coluna HOSTNAME
		table.getColumnModel().getColumn(8).setResizable(false);
		table.getColumnModel().getColumn(8).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(9).setPreferredWidth(75); //coluna FABRICANTE
		table.getColumnModel().getColumn(9).setResizable(false);
		table.getColumnModel().getColumn(9).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(10).setPreferredWidth(100); //coluna MODELO
		table.getColumnModel().getColumn(10).setResizable(false);
		table.getColumnModel().getColumn(10).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(11).setPreferredWidth(80); //coluna SERVICE TAG
		table.getColumnModel().getColumn(11).setResizable(false);
		table.getColumnModel().getColumn(11).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(12).setPreferredWidth(95); //coluna DATA AQUISIÇÃO
		table.getColumnModel().getColumn(12).setResizable(false);
		table.getColumnModel().getColumn(12).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(13).setPreferredWidth(245); //coluna CPU
		table.getColumnModel().getColumn(13).setResizable(false);
		table.getColumnModel().getColumn(13).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(14).setPreferredWidth(90); //coluna STORAGE
		table.getColumnModel().getColumn(14).setResizable(false);
		table.getColumnModel().getColumn(14).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(15).setPreferredWidth(90); //coluna MEMORIA
		table.getColumnModel().getColumn(15).setResizable(false);
		table.getColumnModel().getColumn(15).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(16).setPreferredWidth(160); //coluna TECNICO
		table.getColumnModel().getColumn(16).setResizable(false);
		table.getColumnModel().getColumn(16).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(17).setPreferredWidth(80); //coluna OBSERVAÇÕES
		table.getColumnModel().getColumn(17).setResizable(false);
		table.getColumnModel().getColumn(17).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(18).setPreferredWidth(60); //coluna STATUS
		table.getColumnModel().getColumn(18).setResizable(false);
		table.getColumnModel().getColumn(18).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));

		// Nao vai reodernar os nomes e titulos do cabeçalho da tabela
		table.getTableHeader().setReorderingAllowed(false);

		// Nao permite aumentar na tabela as colunas
		table.setAutoResizeMode(table.AUTO_RESIZE_OFF);

		// Faz com que o usuario selecione um dado na tabela POR VEZ
		table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

		// Realiza a configuracao de fontes
		table.getTableHeader().setFont(new Font("Dialog", Font.BOLD | Font.ITALIC, 10));
		table.setFont(new Font("Dialog", Font.PLAIN, 10));
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

				// Restringir o usuário para selecionar arquivos de todos os tipos
				fc.setAcceptAllFileFilterUsed(false);

				// Coloca um titulo para a janela de dialogo
				fc.setDialogTitle("Selecione um arquivo .txt");

				// Habilita para o user escolher apenas arquivos do tipo: txt
				FileNameExtensionFilter restrict = new FileNameExtensionFilter("Somente arquivos .xlsx", "xlsx");
				fc.addChoosableFileFilter(restrict);
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
				XSSFWorkbook work;
				XSSFSheet sheet;
				XSSFRow row;
				
				//Limpeza dos dados salvos nas listas é como se fosse um F5
				laudo.clear();
				nomeSolicitante.clear();
				usuario.clear();
				centroCusto.clear();
				item.clear();
				qtd.clear();
				ativo.clear();
				dispositivo.clear();
				hostname.clear();
				fabricante.clear();
				modelo.clear();
				serviceTag.clear();
				dataAquisicao.clear();
				cpu.clear();
				storage.clear();
				memoria.clear();
				tecnico.clear();
				observacao.clear();
				status.clear();
				colunas.clear();
				
				
				try {
					int i = 0; //primeira linha
					int pos = 0;
					
					// Abre o arquivo excel
					work = new XSSFWorkbook(new FileInputStream(pathfile));

					// Junta todas as planilhas deste arquivo.
					// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
					sheet = work.getSheetAt(0);

					// Linha de referencia (selecionar a partir de uma determinada linha...)
					row = sheet.getRow(i);

					// Varredura dos cabeçalhos/colunas
					while (row != null) {
						if (row.getCell(pos).getStringCellValue() != null) {
							colunas.add(row.getCell(pos).toString());
							pos++;
						} else {
							work.close();
							break;
						}
					}

					// Chegou na ultima linha que possui conteudo da planilha..
				} catch (FileNotFoundException e1) {
					JOptionPane.showConfirmDialog(null, "Arquivo Excel não encontrado!\nErro: " + e1);
				} catch (NullPointerException e2) {
					System.out.println("Debug error: " + e2 + " Não há mais linhas com conteúdo no excel...");
				} catch (IOException e3) {
					JOptionPane.showConfirmDialog(null, "Arquivo invalido!\nErro: " + e3);
				}
				
				
				try {
					int i = 1; // varredura a partir da segunda linha ~ ignora cabeçalho
					Date date;
					String dataFormatada;

					// Abre o arquivo excel
					work = new XSSFWorkbook(new FileInputStream(pathfile));

					// Junta todas as planilhas deste arquivo.
					// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
					sheet = work.getSheetAt(0);

					// Linha de referencia (selecionar a partir de uma determinada linha...)
					row = sheet.getRow(i);
					
					// Copia as informações das linhas da planilhas para arrayslists.
					while ((row = sheet.getRow(i)) != null) {
						if (row.getCell(0).getNumericCellValue() != 0 && row.getCell(1).getStringCellValue() != null) {
							laudo.add((int)row.getCell(0).getNumericCellValue() + "");
							nomeSolicitante.add(row.getCell(1).getStringCellValue());
							usuario.add(row.getCell(2).getStringCellValue());
							centroCusto.add(row.getCell(3).getStringCellValue());
							item.add(row.getCell(4).getStringCellValue());
							qtd.add((int)row.getCell(5).getNumericCellValue() + "");
							ativo.add(row.getCell(6).getStringCellValue());
							dispositivo.add(row.getCell(7).getStringCellValue());
							hostname.add(row.getCell(8).getStringCellValue());
							fabricante.add(row.getCell(9).getStringCellValue());
							modelo.add(row.getCell(10).getStringCellValue());
							serviceTag.add(row.getCell(11).getStringCellValue());
							
							// Data -> dataAquisicao
							date = row.getCell(12).getDateCellValue();
							SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
							dataFormatada = sdf.format(date);
							dataAquisicao.add(dataFormatada);

							cpu.add(row.getCell(13).getStringCellValue());
							storage.add(row.getCell(14).getStringCellValue());
							memoria.add((int)row.getCell(15).getNumericCellValue() + "");
							tecnico.add(row.getCell(16).getStringCellValue());
							observacao.add(row.getCell(17).getStringCellValue());
							status.add(row.getCell(18).getStringCellValue());
							
							i++;
						} else {
							work.close();
							break;
						}
					}

					// Chegou na ultima linha que possui conteudo da planilha..
				} catch (FileNotFoundException e1) {
					JOptionPane.showConfirmDialog(null, "Arquivo Excel não encontrado!\nErro: " + e1);
				} catch (NullPointerException e2) {
					System.out.println("Debug error: " + e2 + " Não há mais linhas com conteúdo no excel...");
				} catch (IOException e3) {
					JOptionPane.showConfirmDialog(null, "Arquivo invalido!\nErro: " + e3);
				}
				
				
				// remove os espaços que estão ANTES dos textos contidos nas celulas da planilha
				laudo = removeEspacos(laudo);
				nomeSolicitante = removeEspacos(nomeSolicitante);
				usuario = removeEspacos(usuario);
				centroCusto = removeEspacos(centroCusto);
				item = removeEspacos(item);
				qtd = removeEspacos(qtd);
				ativo = removeEspacos(ativo);
				dispositivo = removeEspacos(dispositivo);
				hostname = removeEspacos(hostname);
				fabricante = removeEspacos(fabricante);
				modelo = removeEspacos(modelo);
				serviceTag = removeEspacos(serviceTag);
				dataAquisicao = removeEspacos(dataAquisicao);
				cpu = removeEspacos(cpu);
				storage = removeEspacos(storage);
				memoria = removeEspacos(memoria);
				tecnico = removeEspacos(tecnico);
				observacao = removeEspacos(observacao);
				status = removeEspacos(status);
				
				// Preenche a tabela
				preencherTabelaProprietario();			
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
		
		JButton btnAddLinha = new JButton("Adicionar linha");
		btnAddLinha.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				table.addRowSelectionInterval(0, 2);
			}
		});
		btnAddLinha.setBounds(265, 454, 183, 23);
		contentPane.add(btnAddLinha);
	}

	public ArrayList<String> removeEspacos(ArrayList<String> lista) {
		ArrayList<String> listaSemEspacos = new ArrayList<String>();

		for (int i = 0; i < lista.size(); i++) {
			listaSemEspacos.add(lista.get(i).trim());
		}
		return listaSemEspacos;
	}
}
