//24-09-2023
package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImageOp;
import java.awt.image.ConvolveOp;
import java.awt.image.Kernel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Timer;
import java.util.TimerTask;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.TableCellRenderer;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
//XSSF = (XML SpreadSheet Format) – Used to reading and writting Open Office XML (XLSX) format files.   
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.formdev.flatlaf.FlatIntelliJLaf;
import javax.swing.ImageIcon;
import javax.swing.InputVerifier;

public class Principal extends JFrame {
	// Variaveis Locais
	private static final long serialVersionUID = 6391163855934589017L;
	private JPanel contentPane;
	private JTextField jlocal;
	private JTextField jtitulo;
	static JTable table;
	private static int alinhamento = SwingConstants.LEFT;
	private File pathfileExcel;
	static File pathfileWord;
	private JButton btnPreencher;
	private JButton btnXLS;
	private JButton btnAddLinha;
	static JButton btnSalvarAlteracoes;
	static JButton btnGerarArquivoPdf;
	static JButton btnEditar;
	public int KEYCODE_ESC = 27;
	static Principal frame;
	private JButton btnRemover;
	private int qtdTemporaria;
	private XSSFWorkbook work;
	private int row = 0;
	private int ultimaLinhaOld = 0;
	private boolean flagAdd = false;
	static boolean flagSaved = true;
	private Image img_help = new ImageIcon(SplashAnimation.class.getResource("/resources/help-icon.png")).getImage()
			.getScaledInstance(20, 20, Image.SCALE_SMOOTH);
	private BufferedImage convolvedImage;	
	
	// Database
	static ArrayList<String> laudo = new ArrayList<String>();
	static ArrayList<String> nomeSolicitante = new ArrayList<String>();
	static ArrayList<String> usuario = new ArrayList<String>();
	static ArrayList<String> centroCusto = new ArrayList<String>();
	static ArrayList<String> item = new ArrayList<String>();
	static ArrayList<String> qtd = new ArrayList<String>();
	static ArrayList<String> ativo = new ArrayList<String>();
	static ArrayList<String> dispositivo = new ArrayList<String>();
	static ArrayList<String> hostname = new ArrayList<String>();
	static ArrayList<String> fabricante = new ArrayList<String>();
	static ArrayList<String> modelo = new ArrayList<String>();
	static ArrayList<String> serviceTag = new ArrayList<String>();
	static ArrayList<String> dataAquisicao = new ArrayList<String>();
	static ArrayList<String> cpu = new ArrayList<String>();
	static ArrayList<String> storage = new ArrayList<String>();
	static ArrayList<String> memoria = new ArrayList<String>();
	static ArrayList<String> tecnico = new ArrayList<String>();
	static ArrayList<String> observacao = new ArrayList<String>();
	static ArrayList<String> status = new ArrayList<String>();
	static ArrayList<String> colunas = new ArrayList<String>();
	static ArrayList<Object[]> dados = new ArrayList<>();
	static ArrayList<String> storageType = new ArrayList<String>();
	static ArrayList<String> storageValue = new ArrayList<String>();
	static String[] storageSpliter;
	private JLabel lblHelp;
	private JButton btnFechar;

	public static class HorizontalAlignmentHeaderRenderer implements TableCellRenderer {
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
	public static void preencherTabelaProprietario() {
		dados.clear();
		try {
			for (int i = 0; i < laudo.size(); i++) {
				dados.add(new Object[] { (" " + laudo.get(i)), (" " + nomeSolicitante.get(i)), (" " + usuario.get(i)),
						(" " + centroCusto.get(i)), (" " + item.get(i)), (" " + qtd.get(i)), (" " + ativo.get(i)),
						(" " + dispositivo.get(i)), (" " + hostname.get(i)), (" " + fabricante.get(i)),
						(" " + modelo.get(i)), (" " + serviceTag.get(i)), (" " + dataAquisicao.get(i)),
						(" " + cpu.get(i)), (" " + storage.get(i)), (" " + memoria.get(i)), (" " + tecnico.get(i)),
						(" " + observacao.get(i)), (" " + status.get(i)) });
			}
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "Erro fatal na JTable: " + e);
		}

		// ModeloTabela mode = new ModeloTabela(dados,Colunas);
		ModeloTabela modelo = new ModeloTabela(dados, colunas);
		table.setModel(modelo);
		table.setEnabled(true);

		// Nao deixa a aumentar a largura das colunas da tabela usando o mouse e realiza
		// os alinhamentos das colunas e linhas!
		table.getColumnModel().getColumn(0).setPreferredWidth(50); // coluna LAUDO
		table.getColumnModel().getColumn(0).setResizable(true);
		table.getColumnModel().getColumn(0).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(1).setPreferredWidth(170); // coluna NOME SOLICITANTE
		table.getColumnModel().getColumn(1).setResizable(true);
		table.getColumnModel().getColumn(1).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(2).setPreferredWidth(65); // coluna USUARIO
		table.getColumnModel().getColumn(2).setResizable(true);
		table.getColumnModel().getColumn(2).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(3).setPreferredWidth(200); // coluna CENTRO DE CUSTO
		table.getColumnModel().getColumn(3).setResizable(true);
		table.getColumnModel().getColumn(3).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(4).setPreferredWidth(190); // coluna ITEM
		table.getColumnModel().getColumn(4).setResizable(true);
		table.getColumnModel().getColumn(4).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(5).setPreferredWidth(30); // coluna QTD
		table.getColumnModel().getColumn(5).setResizable(true);
		table.getColumnModel().getColumn(5).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(6).setPreferredWidth(95); // coluna ATIVO
		table.getColumnModel().getColumn(6).setResizable(true);
		table.getColumnModel().getColumn(6).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(7).setPreferredWidth(80); // coluna DISPOSITIVO
		table.getColumnModel().getColumn(7).setResizable(true);
		table.getColumnModel().getColumn(7).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(8).setPreferredWidth(85); // coluna HOSTNAME
		table.getColumnModel().getColumn(8).setResizable(true);
		table.getColumnModel().getColumn(8).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(9).setPreferredWidth(75); // coluna FABRICANTE
		table.getColumnModel().getColumn(9).setResizable(true);
		table.getColumnModel().getColumn(9).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(10).setPreferredWidth(100); // coluna MODELO
		table.getColumnModel().getColumn(10).setResizable(true);
		table.getColumnModel().getColumn(10).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(11).setPreferredWidth(80); // coluna SERVICE TAG
		table.getColumnModel().getColumn(11).setResizable(true);
		table.getColumnModel().getColumn(11).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(12).setPreferredWidth(95); // coluna DATA AQUISIÇÃO
		table.getColumnModel().getColumn(12).setResizable(true);
		table.getColumnModel().getColumn(12).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(13).setPreferredWidth(245); // coluna CPU
		table.getColumnModel().getColumn(13).setResizable(true);
		table.getColumnModel().getColumn(13).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(14).setPreferredWidth(90); // coluna STORAGE
		table.getColumnModel().getColumn(14).setResizable(true);
		table.getColumnModel().getColumn(14).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(15).setPreferredWidth(90); // coluna MEMORIA
		table.getColumnModel().getColumn(15).setResizable(true);
		table.getColumnModel().getColumn(15).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(16).setPreferredWidth(180); // coluna TECNICO
		table.getColumnModel().getColumn(16).setResizable(true);
		table.getColumnModel().getColumn(16).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(17).setPreferredWidth(120); // coluna OBSERVAÇÕES
		table.getColumnModel().getColumn(17).setResizable(true);
		table.getColumnModel().getColumn(17).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));
		table.getColumnModel().getColumn(18).setPreferredWidth(80); // coluna STATUS
		table.getColumnModel().getColumn(18).setResizable(true);
		table.getColumnModel().getColumn(18).setHeaderRenderer(new HorizontalAlignmentHeaderRenderer(alinhamento));

		// Nao vai reodernar os nomes e titulos do cabeçalho da tabela
		table.getTableHeader().setReorderingAllowed(false);

		// Nao permite aumentar na tabela as colunas
		table.setAutoResizeMode(table.AUTO_RESIZE_OFF);

		// Faz com que o usuario selecione um dado na tabela POR VEZ
		table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

		// Realiza a configuracao de fontes
		table.getTableHeader().setFont(new Font("Dialog", Font.BOLD | Font.ITALIC, 10));
		table.setFont(new Font("Dialog", Font.PLAIN, 10));
	}

	public Principal() {
		setIconImage(
				Toolkit.getDefaultToolkit().getImage(Principal.class.getResource("/resources/icon-ServiceDesk.png")));
		setAutoRequestFocus(false);
		setTitle("E-ServiceDesk application");
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 744, 580);
		setLocationRelativeTo(null);

		JMenuBar menuBar = new JMenuBar();
		setJMenuBar(menuBar);

		JMenu mnNewMenu = new JMenu("Ajuda");
		mnNewMenu.setMnemonic('A');
		menuBar.add(mnNewMenu);

		JMenuItem mntmNewMenuItem = new JMenuItem("Sobre E-ServiceDesk Application...");
		mntmNewMenuItem.setMnemonic('S');
		mntmNewMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_DOWN_MASK));
		mntmNewMenuItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String textoParagrafo = "Desenvolvemos esta aplicação para aprimorar a eficiência de nossa equipe de service desk.\n\n"
				        + "Esta é uma ferramenta de automação criada com fins educacionais, visando aprofundar nosso conhecimento em várias APIs, como APACHE POI, JXL, OpenCSV e docx4j, e aprimorar nossas habilidades de programação em um ambiente profissional.\n\n"
				        + "O projeto foi desenvolvido no Eclipse (versão de 2022-09) e construído com o Maven para realizar tarefas como limpeza, verificação, instalação de pacotes e builds. Isso garantiu que todas as APIs necessárias estivessem disponíveis ao alternar entre ambientes domésticos e empresariais, facilitando a colaboração por meio do GitHub, onde você pode acessar nosso perfil e o repositório deste projeto para obter mais informações.\n\n"
				        + "A aplicação oferece uma interface gráfica dinâmica. Ao abri-la, o usuário pode escolher uma planilha específica para manipulação, eliminando a necessidade de usar o Excel. O código foi desenvolvido sob medida para esse tipo de planilha, considerando formatações e quantidade de colunas.\n\n"
				        + "Ao selecionar a planilha, ela é exibida em uma JTable, permitindo a edição como se fosse uma planilha Excel. Além disso, ao clicar em uma linha, você terá opções de edição, e ao clicar duas vezes em uma linha, uma janela lateral abrirá para edição dos valores selecionados. O menu \"Ferramentas\" oferece utilidades adicionais para explorar.\n\n"
				        + "Após realizar as edições desejadas, o usuário pode selecionar uma ou várias linhas e gerar um arquivo PDF. O programa automatiza a criação de um documento MS Word formatado com campos de texto da empresa, servindo como backup. Após a conclusão, o programa converterá o arquivo Word em PDF e permitirá visualizá-lo ou acessar a pasta de destino.\n\n"
				        + "Atenciosamente,\nGustavo Borges.";

				
				 // Crie um JDialog
				JDialog dialogSobre = new JDialog(frame, "Sobre - Pressione ESC para fechar a janela", true);
	            dialogSobre.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
	            // Crie um JTextArea com o texto
	            JTextArea textAreaSobre = new JTextArea(textoParagrafo);

	            textAreaSobre.setWrapStyleWord(true);
	            textAreaSobre.setLineWrap(true);
	            textAreaSobre.setCaretPosition(0); // Posiciona o cursor no início do texto (opcional)
	            textAreaSobre.setAlignmentX(JTextArea.LEFT_ALIGNMENT); // Alinhe à esquerda para justificar
	            textAreaSobre.setAlignmentY(JTextArea.CENTER_ALIGNMENT);
	            textAreaSobre.setEditable(false);
	           // textAreaSobre.setEnabled(false);
	       
	            // Crie um JScrollPane para permitir a rolagem
	            JScrollPane scrollPane = new JScrollPane(textAreaSobre);

	            // Adicione o JScrollPane ao JDialog
	            dialogSobre.getContentPane().add(scrollPane, BorderLayout.CENTER);

	            // Defina o tamanho do JDialog
	            dialogSobre.setSize(560, 520);

	            // Defina a posição do JDialog (centro da tela principal)
	            dialogSobre.setLocationRelativeTo(frame);
	            
	            // Define que a janela nao seja redimensionada manualmente
	            dialogSobre.setResizable(false);
	            // Adicione um KeyListener ao textAreaSobre
	    		dialogSobre.addKeyListener(new KeyAdapter() {
	                @Override
	                public void keyPressed(KeyEvent e) {
	                    if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
	             
	                        dialogSobre.dispose(); // Feche o diálogo ao pressionar "Esc"
	                    }
	                }
	            });
	    
	            // Tornar o JDialog visível
	            dialogSobre.setVisible(true);
	            
	            textAreaSobre.setFocusable(false);
	    		
				// txtpnAplicaoDesenvolvidaPara.setCaretPosition(0);// Sobe para cima a barra de
				// rolagem vertical\
				jlocal.setOpaque(false);

			}
		});


		JMenuItem mntmNewMenuItem_1 = new JMenuItem("Repositorio deste projeto...");
		mntmNewMenuItem_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(null, new MensagemComLink(
						"Conversor XLSX-PDF<br/>Versão V0.2.11.0 <br/><br/>Autor: Gustavo Borges Peres da Silva<br/><a href=\"https://github.com/GustavoBorges13/Conversor_XLSX-PDF\">https://github.com/GustavoBorges13/Conversor_XLSX-PDF</a>"),
						"Informações adicionais", JOptionPane.INFORMATION_MESSAGE);
			}
		});
		mntmNewMenuItem_1.setMnemonic('R');
		mntmNewMenuItem_1.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_R, InputEvent.CTRL_DOWN_MASK));
		mnNewMenu.add(mntmNewMenuItem_1);
		mnNewMenu.add(mntmNewMenuItem);

		JMenu mnNewMenu_2 = new JMenu("Ferramentas");
		menuBar.add(mnNewMenu_2);

		JMenuItem mntmNewMenuItem_2 = new JMenuItem("Opções...");
		mntmNewMenuItem_2.setMnemonic('O');
		mntmNewMenuItem_2.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_DOWN_MASK));
		mntmNewMenuItem_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Opcoes opcoesDialog = new Opcoes();
				opcoesDialog.setVisible(true);
			}
		});
		mnNewMenu_2.add(mntmNewMenuItem_2);
		mntmNewMenuItem_2.setMnemonic('T');

		contentPane = new JPanel();
		contentPane.setOpaque(false);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		SimpleAttributeSet attr = new SimpleAttributeSet();
		attr.copyAttributes();
		StyleConstants.setAlignment(attr, StyleConstants.ALIGN_JUSTIFIED);

		
		jlocal = new JTextField();
		jlocal.setFocusable(true);
		jlocal.setBackground(new Color(240, 240, 240));
		jlocal.setBounds(39, 49, 521, 39);
		jlocal.setForeground(new Color(100, 100, 100));
		contentPane.add(jlocal);
		jlocal.setColumns(10);
		jlocal.setOpaque(true);
		String maskTextXLSX = "Clique na lupa para procurar a planilha...";
		jlocal.setText(maskTextXLSX);

		jlocal.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (jlocal.getText().equals(maskTextXLSX)) {
					jlocal.setText("");
				} else {
					jlocal.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (jlocal.getText().equals("")) {
					jlocal.setText(maskTextXLSX);
				}

			}
		});
		btnXLS = new JButton("");
		Image img_find = new ImageIcon(SplashAnimation.class.getResource("/resources/Find.png")).getImage()
				.getScaledInstance(25, 25, Image.SCALE_SMOOTH);
		btnXLS.setIcon(new ImageIcon(img_find));
		btnXLS.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnXLS.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});

		// BUTTON CHOOSER FILE
		btnXLS.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {
				// File diretorioInicial = new File("\\\\fscatorg01\\...");
				String userHome = System.getProperty("user.home");
				@SuppressWarnings("unused")
				String diretorioInicial = userHome + File.separator + "Downloads";
				JFileChooser fc = new JFileChooser(
						"\\\\fscatorg01\\Fileserver\\Financeiro\\Sistemas\\321 - Sistemas - CAT\\Service Desk\\Documentos\\Laudos Técnicos\\Planilha Laudos");
				fc.setPreferredSize(new Dimension(700, 400));

				// Restringir o usuário para selecionar arquivos de todos os tipos
				fc.setAcceptAllFileFilterUsed(false);

				// Coloca um titulo para a janela de dialogo
				fc.setDialogTitle("Selecione um arquivo .xlsx");

				// Habilita para o user escolher apenas arquivos do tipo: txt
				FileNameExtensionFilter restrict = new FileNameExtensionFilter("Somente arquivos .xlsx", "xlsx");
				fc.addChoosableFileFilter(restrict);
				fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

				int resultado = fc.showOpenDialog(null);

				if (resultado == JFileChooser.CANCEL_OPTION) {
					// Limpa selecoes
					table.clearSelection();
					requestFocus();

					btnFechar.setEnabled(false);
					btnPreencher.setEnabled(false);
				} else {
					pathfileExcel = fc.getSelectedFile();
					jlocal.setText(pathfileExcel.toString().trim());
					jtitulo.setText(fc.getSelectedFile().getName());
					btnPreencher.setEnabled(true); // Habilita o botao

					// Limpa selecoes
					table.clearSelection();
					requestFocus();

					// Limpa tabela se estiver aberta antes
					limparArrayList();
					table.createDefaultColumnsFromModel();

					btnFechar.setEnabled(false);
				}

				// Habilita os botoes auxiliares para controlar a planilha
				btnAddLinha.setEnabled(false);
				btnEditar.setEnabled(false);
				btnRemover.setEnabled(false);
				btnSalvarAlteracoes.setEnabled(false);
				btnGerarArquivoPdf.setEnabled(false);
				btnXLS.setEnabled(true);
			}
		});
		btnXLS.setBounds(592, 49, 47, 39);
		contentPane.add(btnXLS);

		// BUTTON FILL
		btnPreencher = new JButton("Preencher");
		btnPreencher.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnPreencher.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnPreencher.setFocusable(false);
		btnPreencher.setEnabled(false);

		// Button fill/preencher
		btnPreencher.addActionListener(new ActionListener() {
			@SuppressWarnings("unused")
			public void actionPerformed(ActionEvent e) throws NullPointerException {
				btnSalvarAlteracoes.setEnabled(false);
				flagSaved = true;
				work = null;
				XSSFSheet sheet;
				XSSFRow row;

				// Limpeza dos dados salvos nas listas é como se fosse um F5
				limparArrayList();

				try {
					int i = 0; // primeira linha
					int pos = 0;

					// Abre o arquivo excel
					work = new XSSFWorkbook(new FileInputStream(pathfileExcel));
					work.setForceFormulaRecalculation(true);
					// Junta todas as planilhas deste arquivo.
					// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
					sheet = work.getSheetAt(0);

					// Linha de referencia (selecionar a partir de uma determinada linha...)
					row = sheet.getRow(i);

					// Varredura dos cabeçalhos/colunas
					while (row != null) {
						if (!row.getCell(pos).getStringCellValue().isEmpty()) {
							if (row.getCell(pos).getStringCellValue().equals("STATUS")) {
								colunas.add(row.getCell(pos).toString());
								break;
							} else {
								colunas.add(row.getCell(pos).toString());
								pos++;
							}
						}
					}
					work.close();
					// Chegou na ultima linha que possui conteudo da planilha..
				} catch (FileNotFoundException e1) {
					JOptionPane.showMessageDialog(null, "Arquivo Excel não encontrado!\nErro: " + e1, "Erro",
							JOptionPane.ERROR_MESSAGE);
				} catch (IOException e3) {
					JOptionPane.showMessageDialog(null, "Arquivo invalido!\nErro: " + e3, "Erro",
							JOptionPane.ERROR_MESSAGE);
				} finally {
					try {
						int i = 1; // varredura a partir da segunda linha ~ ignora cabeçalho
						int pos = 0; // valor default

						// Abre o arquivo excel
						work = new XSSFWorkbook(new FileInputStream(pathfileExcel));
						work.setForceFormulaRecalculation(true);
						// Junta todas as planilhas deste arquivo.
						// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
						sheet = work.getSheetAt(0);

						// Linha de referencia (selecionar a partir de uma determinada linha...)
						Cell cellSelected;
						if (colunas.size() == 19) {
							// Copia as informações das linhas da planilhas para arrayslists.
							while (((row = sheet.getRow(i)) != null && row.getCell(0).getCellType() != CellType.BLANK
									&& !((int) row.getCell(0).getNumericCellValue() + "").isEmpty())
									|| ((row = sheet.getRow(i)) != null
											&& row.getCell(1).getCellType() != CellType.BLANK
											&& !row.getCell(1).getStringCellValue().isEmpty())
									|| ((row = sheet.getRow(i)) != null
											&& row.getCell(2).getCellType() != CellType.BLANK
											&& !row.getCell(2).getStringCellValue().isEmpty())) {

								storageSpliter = null; // limpa o vetor auxiliar

								// Criando o DataBase Local (Armazenando os valores em Arraylists)
								if (row.getCell(0).getCellType() == CellType.STRING)
									laudo.add(row.getCell(0).getStringCellValue());
								else if (row.getCell(0).getCellType() == CellType.NUMERIC)
									laudo.add((int) row.getCell(0).getNumericCellValue() + "");
								else
									laudo.add("");
								nomeSolicitante.add(row.getCell(1).getStringCellValue());
								usuario.add(row.getCell(2).getStringCellValue());
								centroCusto.add(row.getCell(3).getStringCellValue());
								item.add(row.getCell(4).getStringCellValue());
								if (row.getCell(5).getCellType() == CellType.STRING)
									qtd.add(row.getCell(5).getStringCellValue());
								else
									qtd.add((int) row.getCell(5).getNumericCellValue() + "");
								ativo.add(row.getCell(6).getStringCellValue());
								dispositivo.add(row.getCell(7).getStringCellValue());
								hostname.add(row.getCell(8).getStringCellValue());
								fabricante.add(row.getCell(9).getStringCellValue());
								modelo.add(row.getCell(10).getStringCellValue());
								serviceTag.add(row.getCell(11).getStringCellValue());

								// Tratamento de exeptions
								try {
									// Tentativa de salvar o valor da celular (DATE) na lista do tipo String
									if (DateUtil.isCellDateFormatted(row.getCell(12))) {
										Date date = row.getCell(12).getDateCellValue();
										SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
										String formattedDate = sdf2.format(date);
										dataAquisicao.add(formattedDate);
										// JOptionPane.showMessageDialog(null, "DateString -> " + date + "\nFormatted
										// Date -> " + formattedDate);

										// Caso o valor da celula não seja no formato DATE, verificar se é do tipo
										// numero
									} else if (row.getCell(12).getCellType() == CellType.NUMERIC) {
										String dateString = row.getCell(12).getNumericCellValue() + "";
										try {
											SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

											Date date = sdf.parse(dateString);
											SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
											String formattedDate = sdf2.format(date);
											dataAquisicao.add(formattedDate);
										} catch (ParseException e1) {
											System.out.println("Debug date Formatacao-> " + e1);
										}
									}

									// Erro em caso de nâo ser reconhecido o tipo numerico
								} catch (NullPointerException e1) {
									String dateString = row.getCell(12).getStringCellValue();
									dataAquisicao.add(dateString);
								} catch (IllegalStateException e2) {
									System.out.println("Debug Date Type -> " + e2);
								} finally {
									// Verifica se o tipo é string
									if (row.getCell(12).getCellType() == CellType.STRING) {
										String dateString = row.getCell(12).getStringCellValue();
										if (dateString.equals("N/A") || dateString.equals("N/D")
												|| (dateString = dateString.trim()).equals("")) {
											dataAquisicao.add(dateString);
										} else {
											try {
												SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

												Date date = sdf.parse(dateString);
												SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
												String formattedDate = sdf2.format(date);
												dataAquisicao.add(formattedDate);
												// JOptionPane.showMessageDialog(null, "DateString -> " + dateString +
												// "\nDate -> " + date + "\nFormatted Date -> " + formattedDate);

											} catch (ParseException e2) {
												System.out.println("Debug date -> " + e2);
											}
										}
									}
								}

								cpu.add(row.getCell(13).getStringCellValue());
								if ((row.getCell(14).getStringCellValue()) != null
										&& !(row.getCell(14).getStringCellValue()).isEmpty()) {
									storage.add(row.getCell(14).getStringCellValue());
									storage = removeEspacos(storage);
									storageSpliter = storage.get(i - 1).split(" ");
									storageValue.add(storageSpliter[pos]);
									storageType.add(storageSpliter[pos + 1]);
								}
								if (row.getCell(15).getCellType() == CellType.STRING)
									memoria.add(row.getCell(15).getStringCellValue());
								else
									memoria.add((int) row.getCell(15).getNumericCellValue() + "");
								tecnico.add(row.getCell(16).getStringCellValue());

								// Observacao
								Cell cell = row.getCell(17);
								if (cell == null || cell.getStringCellValue().isEmpty()) {
									observacao.add("");
								} else {
									observacao.add(cell.getStringCellValue());
								}

								// Status
								cell = row.getCell(18);
								if (cell == null || cell.getStringCellValue().isEmpty()) {
									status.add("");
								} else {
									status.add(cell.getStringCellValue());
								}

								i++;
							}
							qtdTemporaria = laudo.size();
							work.close();

						} else {
							work.close();
							throw new java.lang.ArrayIndexOutOfBoundsException();
						}
					} catch (IOException e2) {
						JOptionPane.showMessageDialog(null, "Arquivo invalido!\nErro: " + e2, "Erro",
								JOptionPane.ERROR_MESSAGE);
					} catch (java.lang.ArrayIndexOutOfBoundsException e3) {
						JOptionPane.showMessageDialog(null,
								"Esta não é a planilha que utilizamos em 2023.\nPor favor, abra a planilha com o modelo padrão utilizado pois,"
										+ " este codigo foi feito especificamente para ser utilizado com esse tipo de planilha devido as formatações"
										+ " e quantidade de colunas.\nErro: " + e3,
								"Aviso", JOptionPane.WARNING_MESSAGE);
					} catch (java.lang.NullPointerException e4) {
						JOptionPane.showMessageDialog(null,
								"Está é uma planilha nova! Possivelmente existem campos vazios, por favor, insira novos laudos.\n Exception: "
										+ e4,
								"Aviso", JOptionPane.WARNING_MESSAGE);
					}
				}

				// remove os espaços/pontos que estão ANTES dos textos contidos nas celulas da
				// planilha
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
				fabricante = removePontuacoes(fabricante);
				modelo = removeEspacos(modelo);
				serviceTag = removeEspacos(serviceTag);
				dataAquisicao = removeEspacos(dataAquisicao);
				cpu = removeEspacos(cpu);
				// storage = removeEspacos(storage); //foi usado no looping ^-^
				memoria = removeEspacos(memoria);
				tecnico = removeEspacos(tecnico);
				observacao = removeEspacos(observacao);
				status = removeEspacos(status);

				dados.clear();

				// Preenche a tabela
				preencherTabelaProprietario();

				// Anota quantas linhas de dados tem
				ultimaLinhaOld = table.getRowCount();

				// Habilita os botoes auxiliares para controlar a planilha
				btnAddLinha.setEnabled(true);
				btnEditar.setEnabled(false);
				btnRemover.setEnabled(false);
				btnFechar.setEnabled(true);
				try {
					work.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});

		btnPreencher.setBounds(591, 131, 105, 39);
		contentPane.add(btnPreencher);

		jtitulo = new JTextField();
		jtitulo.setEditable(false);
		jtitulo.setDisabledTextColor(new Color(0, 0, 0));
		jtitulo.setColumns(10);
		jtitulo.setBounds(39, 131, 521, 39);
		contentPane.add(jtitulo);

		JLabel lblTituloPlanilha = new JLabel("Nome da planilha");
		lblTituloPlanilha.setBounds(39, 106, 105, 14);
		contentPane.add(lblTituloPlanilha);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(39, 195, 657, 279);
		contentPane.add(scrollPane);

		// TABLE - TABELA
		table = new JTable();
		table.setAutoscrolls(false);
		table.addMouseListener(new MouseAdapter() {
			boolean isAlreadyOneClick = false;

			@Override
			public void mouseEntered(MouseEvent e) {
				if (table.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}

			@Override
			public void mouseClicked(MouseEvent arg0) {
				if (isAlreadyOneClick) {
					int linhaSelecionada = table.getSelectedRow();
					// int[] selectedRows = table.getSelectedRows();
					EditarPlanilha frame = new EditarPlanilha();
					int toleranciaHD_SSD = 0;
					int pos = 0; // valor default

					// Foco total para a nova janela aberta
					frame.setVisible(true);
					frame.requestFocus();
					frame.setResizable(false);

					// Deixa a janela principal desativada
					Principal.this.setEnabled(false);

					// Evento pra verificar se a janela de edicao foi fechada
					frame.addWindowListener(new WindowAdapter() {
						@Override
						public void windowClosed(WindowEvent e) {
							if (!EditarPlanilha.finalizado) {
								int result = JOptionPane.showConfirmDialog(frame,
										"Você deseja realmente cancelar as operações e fechar a janela?", "Confirmação",
										JOptionPane.YES_NO_OPTION);
								if (result == JOptionPane.YES_OPTION) {
									frame.setVisible(false);
									int selectedRow = table.getSelectedRow();
									if (selectedRow != -1 && table.getValueAt(selectedRow, 0) == "") {
										btnRemover.doClick();
									}

									Principal.this.setEnabled(true);

									// Habilita/desabilita botoes
									btnEditar.setEnabled(false);
									btnRemover.setEnabled(false);
									table.clearSelection();
									contentPane.requestFocus();
									btnGerarArquivoPdf.setEnabled(false);

									// flag
									flagSaved = true;
									flagAdd = false;
									// frame.setVisible(false);
								} else {
									frame.setVisible(true);
								}
							} else {
								Principal.this.setEnabled(true);

								// Habilita/desabilita botoes
								btnEditar.setEnabled(false);
								btnRemover.setEnabled(false);
								table.clearSelection();
								contentPane.requestFocus();
								btnGerarArquivoPdf.setEnabled(false);

								// Atualiza a tabela
								// Principal.limpaListas();
								preencherTabelaProprietario();
								table.updateUI();
								table.clearSelection();
								table.requestFocus();

								// table.setRowSelectionInterval(table.getRowCount()-1,table.getRowCount()-1);
							}
						}

						@Override
						public void windowDeactivated(WindowEvent e) {
							// Focando a janela de gerar laudo se o usuario minimizou tudo clicando na area
							// de trabalho
							frame.requestFocus();
						}
					});

					linhaSelecionada = table.getSelectedRow();

					// Design - Deixar os textos em preto e Transpor dados da tabela para a janela
					// nova
					EditarPlanilha.txtLaudo.setForeground(Color.BLACK);
					EditarPlanilha.txtLaudo.setText(laudo.get(linhaSelecionada));

					EditarPlanilha.txtNomeSolicitante.setForeground(Color.BLACK);
					EditarPlanilha.txtNomeSolicitante.setText(nomeSolicitante.get(linhaSelecionada));

					EditarPlanilha.txtUsuario.setForeground(Color.BLACK);
					EditarPlanilha.txtUsuario.setText(usuario.get(linhaSelecionada));

					EditarPlanilha.txtCentroDeCusto.setForeground(Color.BLACK);
					EditarPlanilha.txtCentroDeCusto.setText(centroCusto.get(linhaSelecionada));

					EditarPlanilha.txtItem.setForeground(Color.BLACK);
					EditarPlanilha.txtItem.setText(item.get(linhaSelecionada));

					for (int i = 0; i <= EditarPlanilha.comboBoxQuantidade.getItemCount() - 1; i++) {
						if (!qtd.get(linhaSelecionada).equals(EditarPlanilha.comboBoxQuantidade.getItemAt(i)))
							EditarPlanilha.comboBoxQuantidade.setForeground(Color.RED);
						else {
							EditarPlanilha.comboBoxQuantidade.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxQuantidade.setSelectedIndex(i);
							break;
						}
					}

					EditarPlanilha.txtAtivo.setForeground(Color.BLACK);
					EditarPlanilha.txtAtivo.setText(ativo.get(linhaSelecionada));

					for (int i = 0; i <= EditarPlanilha.comboBoxDispositivo.getItemCount() - 1; i++) {
						if (!dispositivo.get(linhaSelecionada).equals(EditarPlanilha.comboBoxDispositivo.getItemAt(i)))
							EditarPlanilha.comboBoxDispositivo.setForeground(Color.RED);
						else {
							EditarPlanilha.comboBoxDispositivo.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxDispositivo.setSelectedIndex(i);
							break;
						}
					}

					EditarPlanilha.txtHostname.setForeground(Color.BLACK);
					EditarPlanilha.txtHostname.setText(hostname.get(linhaSelecionada));

					for (int i = 0; i <= EditarPlanilha.comboBoxFabricante.getItemCount() - 1; i++) {
						if (!fabricante.get(linhaSelecionada).equals(EditarPlanilha.comboBoxFabricante.getItemAt(i))) {
							if (EditarPlanilha.comboBoxFabricante.getItemAt(i).toString().toUpperCase()
									.contains(fabricante.get(linhaSelecionada).toUpperCase())) {
								EditarPlanilha.comboBoxFabricante.setForeground(Color.BLACK);
								EditarPlanilha.comboBoxFabricante.setSelectedIndex(i);
								break;
							}
							EditarPlanilha.comboBoxFabricante.setForeground(Color.RED);
						} else {
							EditarPlanilha.comboBoxFabricante.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxFabricante.setSelectedIndex(i);
							break;
						}
					}

					EditarPlanilha.txtModelo.setForeground(Color.BLACK);
					EditarPlanilha.txtModelo.setText(modelo.get(linhaSelecionada));

					EditarPlanilha.txtServiceTag.setForeground(Color.BLACK);
					EditarPlanilha.txtServiceTag.setText(serviceTag.get(linhaSelecionada));

					EditarPlanilha.txtDdmmyyyy.setForeground(Color.BLACK);
					EditarPlanilha.txtDdmmyyyy.setText(dataAquisicao.get(linhaSelecionada));

					EditarPlanilha.txtCpu.setForeground(Color.BLACK);
					EditarPlanilha.txtCpu.setText(cpu.get(linhaSelecionada));

					// System.out.println("Type: "+storageType.get(linhaSelecionada)+" Value:
					// "+Integer.parseInt(storageValue.get(linhaSelecionada)));
					// --- HD ---
					if (flagAdd == true && (ultimaLinhaOld <= table.getRowCount())) {
						storageSpliter = storage.get(linhaSelecionada).split(" ");
						storageValue.set(linhaSelecionada, storageSpliter[pos]);
						storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
					}

					if (storageType.get(linhaSelecionada).equals("HD")) {
						toleranciaHD_SSD = 20;
						if (Integer.parseInt(storageValue.get(linhaSelecionada)) >= 0
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 180 + toleranciaHD_SSD) { // Range
																														// 180GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "180 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 180 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 250 + toleranciaHD_SSD) { // Range
																														// 250GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "250 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 250 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 300 + toleranciaHD_SSD) { // Range
																														// 300GB
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "300 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 300 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 500 + toleranciaHD_SSD) { // Range
																														// 500GB
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "500 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 500 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 750 + toleranciaHD_SSD) { // Range
																														// 750GB
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "750 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 750 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 1000 + toleranciaHD_SSD) { // Range
																														// 1000GB
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "1000 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 1000 + toleranciaHD_SSD) { // Range
																														// +1000GB
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "1000 HD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						}

						// --- SSD ---
					} else if (storageType.get(linhaSelecionada).equals("SSD")) {
						toleranciaHD_SSD = 30;
						if (Integer.parseInt(storageValue.get(linhaSelecionada)) >= 0
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 120 + toleranciaHD_SSD) { // Range
																														// 120GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "120 SSD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 120 + toleranciaHD_SSD
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 240 + toleranciaHD_SSD) { // Range
																														// 240GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "240 SSD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						} else if (Integer.parseInt(storageValue.get(linhaSelecionada)) > 120 + toleranciaHD_SSD) { // Range
																													// +240GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "240 SSD"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						}

						// --- SSD-NVMe ---
					} else if (storageType.get(linhaSelecionada).equals("SSD-NVMe")) {
						toleranciaHD_SSD = 50;
						if (Integer.parseInt(storageValue.get(linhaSelecionada)) >= 0
								&& Integer.parseInt(storageValue.get(linhaSelecionada)) <= 256 + toleranciaHD_SSD) { // Range
																														// 256GB
							storageSpliter = null;
							// Atribuindo um novo valor arredondado/aproximado caso o valor encontrado não
							// seja compativel com as nossas amostras disponiveis.
							storage.set(linhaSelecionada, "256 SSD-NVMe"); // Valor novo!!!
							storageSpliter = storage.get(linhaSelecionada).split(" ");
							storageValue.set(linhaSelecionada, storageSpliter[pos]);
							storageType.set(linhaSelecionada, storageSpliter[pos + 1]);
							EditarPlanilha.comboBoxStorage.setForeground(Color.BLACK);
							EditarPlanilha.comboBoxStorage.setSelectedItem(storage.get(linhaSelecionada));

						}
					} else {
						storageSpliter = null;
						// PADRAO
						EditarPlanilha.comboBoxStorage.setForeground(Color.RED);
						EditarPlanilha.comboBoxStorage.setSelectedIndex(0);
					}

					int memoryTemp = Integer.parseInt(memoria.get(linhaSelecionada));
					if (memoryTemp == 0) {
						EditarPlanilha.c.setForeground(Color.RED);
						EditarPlanilha.spinner_memoria.setValue(Integer.parseInt(memoria.get(linhaSelecionada)));

					} else {
						EditarPlanilha.c.setForeground(Color.BLACK);
						EditarPlanilha.spinner_memoria.setValue(Integer.parseInt(memoria.get(linhaSelecionada)));
					}

					EditarPlanilha.txtNomeDoTecnico.setForeground(Color.BLACK);
					EditarPlanilha.txtNomeDoTecnico.setText(tecnico.get(linhaSelecionada));

					EditarPlanilha.txtObservao.setForeground(Color.BLACK);
					EditarPlanilha.txtObservao.setText(observacao.get(linhaSelecionada));

					EditarPlanilha.txtStatus.setForeground(Color.BLACK);
					EditarPlanilha.txtStatus.setText(status.get(linhaSelecionada));

					isAlreadyOneClick = false;
				} else {
					isAlreadyOneClick = true;
					Timer t = new Timer("doubleclickTimer", false);

					// Habilita botoes
					btnEditar.setEnabled(true);
					btnRemover.setEnabled(true);
					btnGerarArquivoPdf.setEnabled(true);

					// Limpeza de selecao
					// table.clearSelection();
					// table.requestDefaultFocus();
					t.schedule(new TimerTask() {
						@Override
						public void run() {
							isAlreadyOneClick = false;
						}
					}, 300);
				}

			}

			@Override
			public void mousePressed(MouseEvent e) {
				int linhaSelecionada = table.getSelectedRow();
				if (table.isRowSelected(linhaSelecionada)) {
					// Habilita botoes
					btnEditar.setEnabled(true);
					btnRemover.setEnabled(true);
					btnGerarArquivoPdf.setEnabled(true);
				}
			}
		});
		scrollPane.setViewportView(table);

		// BUTTO ADD ROW
		btnAddLinha = new JButton("Adicionar");
		btnAddLinha.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnAddLinha.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnAddLinha.setEnabled(false);
		btnAddLinha.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				// Habilita botoes
				btnEditar.setEnabled(true);
				btnRemover.setEnabled(true);
				btnSalvarAlteracoes.setEnabled(true);

				// flag
				flagSaved = false;

				// Adiciona uma linha em branco ao final da tabela
				dados.add(new Object[] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" });
				adicionarArrayList();

				// Atualiza a planilha
				table.updateUI();
				table.requestFocus();
				table.setRowSelectionInterval(table.getRowCount() - 1, table.getRowCount() - 1);
				int selectedRow = table.getSelectedRow();

				if (selectedRow != -1) {
					// Point p = table.getMousePosition();
					MouseEvent me = new MouseEvent(table, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 0, 0,
							1, false);
					table.dispatchEvent(me);
					me = new MouseEvent(table, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 0, 0, 1, false,
							MouseEvent.BUTTON1);
					table.dispatchEvent(me);

				}
			}
		});
		btnAddLinha.setBounds(154, 485, 105, 23);
		contentPane.add(btnAddLinha);

		JLabel lblSelecioneUmaPlanilha = new JLabel("Selecione a planilha de laudo técnico .xlsx");
		lblSelecioneUmaPlanilha.setBounds(39, 24, 393, 14);
		contentPane.add(lblSelecioneUmaPlanilha);

		// BUTTON EDIT
		btnEditar = new JButton("Editar");
		btnEditar.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnEditar.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnEditar.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				int[] selectedRows = table.getSelectedRows();
				if (selectedRows.length != -1 && selectedRows.length <= 1) {

					// Point p = table.getMousePosition();
					MouseEvent me = new MouseEvent(table, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 0, 0,
							1, false);
					table.dispatchEvent(me);
					me = new MouseEvent(table, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 0, 0, 1, false,
							MouseEvent.BUTTON1);
					table.dispatchEvent(me);
				} else {
					JOptionPane
							.showMessageDialog(null,
									"Por favor selecione apenas uma linha por vez para editar.\nVocê tentou editar com "
											+ selectedRows.length + " linhas selecionadas.",
									"Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		btnEditar.setEnabled(false);
		btnEditar.setBounds(39, 485, 105, 23);
		contentPane.add(btnEditar);

		// BUTTON SAVE
		btnSalvarAlteracoes = new JButton("Salvar alterações");
		btnSalvarAlteracoes.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnSalvarAlteracoes.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnSalvarAlteracoes.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int reply = JOptionPane.showConfirmDialog(null,
						"Você quer salvar as alterações realizadas na planilha?", "Salvamento",
						JOptionPane.YES_NO_OPTION);
				if (reply == JOptionPane.YES_OPTION) {
					try {
						// Abra o arquivo xlsx existente
						int g = 1;
						work = new XSSFWorkbook(new FileInputStream(pathfileExcel));
						work.setForceFormulaRecalculation(true);
						// Junta todas as planilhas deste arquivo.
						// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
						// Obtenha a planilha desejada
						XSSFSheet sheet = work.getSheetAt(0);
						XSSFRow row;
						XSSFCell cell = null;

						// Configurando estilo
						XSSFCellStyle textStyle = work.createCellStyle();
						textStyle.setAlignment(HorizontalAlignment.CENTER);
						textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
						textStyle.setBorderTop(BorderStyle.THIN);
						textStyle.setBorderBottom(BorderStyle.THIN);
						textStyle.setBorderLeft(BorderStyle.THIN);
						textStyle.setBorderRight(BorderStyle.THIN);

						// Remove todo o conteudo da planilha para fazer o reset
						if (laudo.size() < qtdTemporaria)
							sheet.copyRows(sheet.getLastRowNum() - qtdTemporaria, sheet.getLastRowNum(), 1,
									new CellCopyPolicy());
						else
							sheet.copyRows(sheet.getLastRowNum() - laudo.size(), sheet.getLastRowNum(), 1,
									new CellCopyPolicy());

						// Calcula g
						while ((row = sheet.getRow(g)) != null && row.getCell(0).getCellType() == CellType.BLANK
								&& row.getCell(1).getCellType() == CellType.BLANK
								&& row.getCell(2).getCellType() == CellType.BLANK
								&& row.getCell(3).getCellType() != CellType.BLANK) {
							g++;
						}
						// SALVANDO NEW DADOS NA PLANILHA!
						escreverPlanilha(cell, sheet, textStyle);

						// Adicionando mais linhas se estiver perto da linha de aviso (AMARELO)
						int limitador = 3;
						if ((g - laudo.size()) <= limitador) {
							for (int i = 0; i <= limitador; i++) {
								addRow(sheet, 1, laudo.size() + 1, textStyle);
							}
						}

						// Salvar alterações na planilha!
						work.write(new FileOutputStream(pathfileExcel));

						work.close(); // Fecha a planilha
						JOptionPane.showMessageDialog(null, "As alterações foram salvas com suceeso!");

						// Atualiza a tabela...
						btnPreencher.doClick();

						// Desabilita botoes
						btnSalvarAlteracoes.setEnabled(false);
						btnEditar.setEnabled(false);
						btnGerarArquivoPdf.setEnabled(false);

						// Alterar flags
						flagAdd = true;

					} catch (IOException e) {
						JOptionPane.showMessageDialog(null, "Erro com o arquivo!\n" + e.getMessage()
								+ "\nPor favor, feche o programa que esteja utilizando o arquivo que você quer salvar.",
								"Erro", JOptionPane.ERROR_MESSAGE);
					} /*
						 * catch (NullPointerException e) { JOptionPane.showMessageDialog(null,
						 * "As alterações foram salvas com suceeso!");
						 * 
						 * // Atualiza a tabela... btnPreencher.doClick();
						 * JOptionPane.showMessageDialog(null, aux); // Desabilita botoes
						 * btnSalvarAlteracoes.setEnabled(false); btnEditar.setEnabled(false);
						 * btnGerarArquivoPdf.setEnabled(false); }
						 */
				} else {
					requestFocus();
				}
			}
		});
		btnSalvarAlteracoes.setEnabled(false);
		btnSalvarAlteracoes.setBounds(384, 485, 138, 23);
		contentPane.add(btnSalvarAlteracoes);

		// BUTTON GENERATE PDF FILE
		btnGerarArquivoPdf = new JButton("Gerar arquivo PDF");
		btnGerarArquivoPdf.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnGerarArquivoPdf.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnGerarArquivoPdf.addActionListener(new ActionListener() {
			@SuppressWarnings("unused")
			public void actionPerformed(ActionEvent arg0) {
				boolean flagContinuacao = false;
				int[] linhasSelecionadas = Principal.table.getSelectedRows();
				String aux = null;

				// Verifica se os vetores têm elementos em comum
				boolean temElementosEmComum = false;
				if (linhasSelecionadas.length > 1) {
					for (int i = 0; i < linhasSelecionadas.length; i++) {
						for (int j = i + 1; j < linhasSelecionadas.length; j++) {
							if (Principal.table.getValueAt(linhasSelecionadas[i], 8)
									.equals(Principal.table.getValueAt(linhasSelecionadas[j], 8))
									|| Principal.table.getValueAt(linhasSelecionadas[i], 8 - 2)
											.equals(Principal.table.getValueAt(linhasSelecionadas[j], 8 - 2))) {
								temElementosEmComum = true;
							} else {
								flagContinuacao = false;
								temElementosEmComum = false;
								break;
							}
						}
						if (temElementosEmComum) {
							flagContinuacao = true;
							break;
						}
					}
				} else {
					flagContinuacao = true;
				}
				if (flagContinuacao && flagSaved) {

					GerarLaudoPDF frame = new GerarLaudoPDF();

					// Foco total para a nova janela aberta
					frame.setVisible(true);
					frame.requestFocus();

					// Deixa a janela principal desativada
					Principal.this.setEnabled(false);

					// Evento pra verificar se a janela de edicao foi fechada
					frame.addWindowListener(new WindowAdapter() {
						@Override
						public void windowClosed(WindowEvent e) {
							if (!GerarLaudoPDF.finalizado) {
								int result = JOptionPane.showConfirmDialog(frame,
										"Você deseja realmente cancelar as operações e fechar a janela?", "Confirmação",
										JOptionPane.YES_NO_OPTION);
								if (result == JOptionPane.YES_OPTION) {
									frame.setVisible(false);
									Principal.this.setEnabled(true);

									// Habilita/desabilita botoes
									btnEditar.setEnabled(false);
									btnRemover.setEnabled(false);
									table.clearSelection();
									contentPane.requestFocus();
									btnGerarArquivoPdf.setEnabled(false);
									GerarLaudoPDF.linkExibicao.clear();
									GerarLaudoPDF.linkEndereco.clear();
								} else {
									frame.setVisible(true);
								}
							} else {
								Principal.this.setEnabled(true);
								btnEditar.setEnabled(false);
								table.requestFocus();
							}
						}

						@Override
						public void windowDeactivated(WindowEvent e) {
							// Focando a janela de gerar laudo se o usuario minimizou tudo clicando na area
							// de trabalho
							frame.requestFocus();
						}

					});
				} else if (!flagSaved) {
					JOptionPane.showMessageDialog(null,
							"Infelizmente existem alterações realizadas na planilha que não foram salvas, por favor, salve as alterações antes continuar.",
							"Aviso", JOptionPane.WARNING_MESSAGE);

					// Salva as alteracoes antes
					btnSalvarAlteracoes.doClick();

				} else {
					JOptionPane.showMessageDialog(null,
							"Por favor, verifique as linhas que você selecionou, se possuem os mesmos ativos.\nPois não é permitido fazer 1 laudo para computadores diferentes.",
							"Erro", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		btnGerarArquivoPdf.setEnabled(false);
		btnGerarArquivoPdf.setBounds(532, 485, 164, 23);
		contentPane.add(btnGerarArquivoPdf);

		// BUTTON REMOVE
		btnRemover = new JButton("Remover");
		btnRemover.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnRemover.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnRemover.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int selectedRow = table.getSelectedRow();
				if (selectedRow != -1) {
					ModeloTabela model = (ModeloTabela) table.getModel();

					// Remove a linha da tabela
					model.removeRow(selectedRow);

					// Limpa a base de dados
					removerArrayList(selectedRow);

					// Atualiza a tabela
					preencherTabelaProprietario();
					table.updateUI();
					table.requestFocus();

					// Desabilita/habilita botoes
					btnEditar.setEnabled(false);
					btnRemover.setEnabled(false);
					btnSalvarAlteracoes.setEnabled(true);
					btnGerarArquivoPdf.setEnabled(false);

					// flag
					flagSaved = false;
				}
			}
		});
		btnRemover.setEnabled(false);
		btnRemover.setBounds(269, 485, 105, 23);
		contentPane.add(btnRemover);

		lblHelp = new JLabel("Atalhos");
		lblHelp.setHorizontalTextPosition(SwingConstants.LEFT);
		lblHelp.setHorizontalAlignment(SwingConstants.RIGHT);
		lblHelp.setIcon(new ImageIcon(img_help));
		lblHelp.setBounds(640, 0, 85, 23);
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
				JOptionPane.showMessageDialog(null, "Atalhos ");
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

		btnFechar = new JButton("");
		btnFechar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				int reply = JOptionPane.showConfirmDialog(null, "Você tem certeza que deseja fechar a planilha?",
						"Fechar planilha", JOptionPane.WARNING_MESSAGE, JOptionPane.YES_NO_OPTION);
				if (reply == JOptionPane.YES_OPTION) {
					work.setForceFormulaRecalculation(true);
					try {
						work.close();
					} catch (IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
					dispose();
					Principal pp = new Principal();
					pp.setVisible(true);
				}
			}
		});
		btnFechar.setEnabled(false);
		Image img_close = new ImageIcon(SplashAnimation.class.getResource("/resources/closeRed.png")).getImage()
				.getScaledInstance(25, 25, Image.SCALE_SMOOTH);
		btnFechar.setIcon(new ImageIcon(img_close));
		btnFechar.setBounds(649, 49, 47, 39);
		btnFechar.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnFechar.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});

		contentPane.add(btnFechar);

		// Copy dados da tabela (atalhos)
		table.addKeyListener(new KeyAdapter() {

			@Override
			public void keyPressed(KeyEvent e) {
				if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_C) {
					row = table.getSelectedRow();
				}
				if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_V) {
					if (row != -1 && table.getValueAt(row, 0) != null) {
						// Atualiza o banco de dados
						copiarArrayList(row);

						// Atualiza a tabela com os novos dados
						preencherTabelaProprietario();
						((ModeloTabela) table.getModel()).fireTableDataChanged();
						table.updateUI();

						// preencherTabelaProprietario();

						// Habilita ou desabilita botoes
						btnEditar.setEnabled(false);
						btnRemover.setEnabled(false);
						btnSalvarAlteracoes.setEnabled(true);
						btnGerarArquivoPdf.setEnabled(false);

						// Flag
						flagSaved = false;
					}
				}
				if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_X) {
					btnRemover.doClick();
				}
				if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_B) {
					btnSalvarAlteracoes.doClick();
				}
				if (e.isControlDown() && e.isShiftDown() && e.getKeyCode() == KeyEvent.VK_N) {
					btnAddLinha.doClick();
				}
				if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_G) {
					btnGerarArquivoPdf.doClick();
				}
				if (e.getKeyCode() == KeyEvent.VK_F5 || (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_Z)) {
					btnPreencher.doClick();
				}
			}
		});
		
		
		
		btnXLS.setFocusable(true);
		btnXLS.requestFocus();
	}

	public ArrayList<String> removeEspacos(ArrayList<String> lista) {
		ArrayList<String> listaSemEspacos = new ArrayList<String>();

		for (int i = 0; i < lista.size(); i++) {
			listaSemEspacos.add(lista.get(i).trim());
		}
		return listaSemEspacos;

	}

	public ArrayList<String> removePontuacoes(ArrayList<String> lista) {
		ArrayList<String> listaSemEspacos = new ArrayList<String>();

		for (int i = 0; i < lista.size(); i++) {
			listaSemEspacos.add(lista.get(i).replace(".", ""));
		}
		return listaSemEspacos;

	}

	public static void copiarArrayList(int linha) {
		dados.add(dados.get(linha));
		laudo.add(laudo.get(linha));
		nomeSolicitante.add(nomeSolicitante.get(linha));
		usuario.add(usuario.get(linha));
		centroCusto.add(centroCusto.get(linha));
		item.add(item.get(linha));
		qtd.add(qtd.get(linha));
		ativo.add(ativo.get(linha));
		dispositivo.add(dispositivo.get(linha));
		hostname.add(hostname.get(linha));
		fabricante.add(fabricante.get(linha));
		modelo.add(modelo.get(linha));
		serviceTag.add(serviceTag.get(linha));
		dataAquisicao.add(dataAquisicao.get(linha));
		cpu.add(cpu.get(linha));
		storage.add(storage.get(linha));
		storageType.add(storageType.get(linha));
		storageValue.add(storageValue.get(linha));
		memoria.add(memoria.get(linha));
		tecnico.add(tecnico.get(linha));
		observacao.add(observacao.get(linha));
		status.add(status.get(linha));
	}

	public static void limparArrayList() {
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
		dados.clear();
		storageSpliter = null;
		storageType.clear();
		storageValue.clear();
	}

	public void adicionarArrayList() {
		laudo.add("");
		nomeSolicitante.add("");
		usuario.add("");
		centroCusto.add("");
		item.add("");
		qtd.add(null);
		ativo.add("");
		dispositivo.add("");
		hostname.add("");
		fabricante.add("");
		modelo.add("");
		serviceTag.add("");
		dataAquisicao.add("");
		cpu.add("");
		storage.add("");
		storageType.add("");
		storageValue.add("");
		memoria.add("0");
		tecnico.add("");
		observacao.add("");
		status.add("");
	}

	public void removerArrayList(int pos) {
		dados.removeIf(n -> (dados.size() == 0));
		laudo.remove(pos);
		nomeSolicitante.remove(pos);
		usuario.remove(pos);
		centroCusto.remove(pos);
		item.remove(pos);
		qtd.remove(pos);
		ativo.remove(pos);
		dispositivo.remove(pos);
		hostname.remove(pos);
		fabricante.remove(pos);
		modelo.remove(pos);
		serviceTag.remove(pos);
		dataAquisicao.remove(pos);
		cpu.remove(pos);
		storage.remove(pos);
		storageType.remove(pos);
		storageValue.remove(pos);
		memoria.remove(pos);
		tecnico.remove(pos);
		observacao.remove(pos);
		status.remove(pos);
	}

	private static void addRow(XSSFSheet worksheet, int sourceRowNum, int destinationRowNum, XSSFCellStyle textStyle) {
		// Get the source / new row
		XSSFRow newRow = worksheet.getRow(destinationRowNum);
		XSSFRow sourceRow = worksheet.getRow(sourceRowNum);
		// If the row exist in destination, push down all rows by 1 else create a new
		// row
		if (newRow != null) {
			worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
		} else {
			newRow = worksheet.createRow(destinationRowNum);
		}

		// Loop through source columns to add to new row
		for (int i = 0; i < colunas.size(); i++) {
			// Grab a copy of the old/new cell
			XSSFCell oldCell = sourceRow.getCell(i);
			XSSFCell newCell = newRow.createCell(i);

			// If the old cell is null jump to next cell
			if (oldCell == null) {
				newCell = null;
				continue;
			}

			newCell.setCellStyle(textStyle);

			// If there is a cell comment, copy
			if (oldCell.getCellComment() != null) {
				newCell.setCellComment(oldCell.getCellComment());
			}

			// If there is a cell hyperlink, copy
			if (oldCell.getHyperlink() != null) {
				newCell.setHyperlink(oldCell.getHyperlink());
			}

			// Set the cell data type
			newCell.setCellType(oldCell.getCellType());
		}
	}

	public static void escreverPlanilha(XSSFCell cell, XSSFSheet sheet, XSSFCellStyle textStyle) {
		int j = 0;
		try {
			// Transcreve os valores do arraylist para as celulas da planilha...
			for (int i = 1; i <= laudo.size(); i++) { // Obtendo a linha desejada
				XSSFRow row = sheet.getRow(i);

				// SALVANDO NEW DADOS NA PLANILHA!
				// Alterando valores das celulas...
				cell = row.getCell(0); // Obtem a celula desejada -> Coluna
				cell.setCellValue(Integer.parseInt(laudo.get(j))); // Alterando o valor das células da planilha
				cell.setCellType(CellType.NUMERIC); // Define o tipo da celula
				cell.setCellStyle(textStyle); // Define o estilo centralizado
				cell = row.getCell(1);
				cell.setCellValue(nomeSolicitante.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(2);
				cell.setCellValue(usuario.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(3);
				cell.setCellValue(centroCusto.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(4);
				cell.setCellValue(item.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(5);
				cell.setCellValue(Integer.parseInt(qtd.get(j)));
				cell.setCellType(CellType.NUMERIC);
				cell.setCellStyle(textStyle);
				cell = row.getCell(6);
				cell.setCellValue(ativo.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(7);
				cell.setCellValue(dispositivo.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(8);
				cell.setCellValue(hostname.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(9);
				cell.setCellValue(fabricante.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(10);
				cell.setCellValue(modelo.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(11);
				cell.setCellValue(serviceTag.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(12);
				cell.setCellValue(dataAquisicao.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(13);
				cell.setCellValue(cpu.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(14);
				cell.setCellValue(storage.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(15);
				cell.setCellValue(Integer.parseInt(memoria.get(j)));
				cell.setCellType(CellType.NUMERIC);
				cell.setCellStyle(textStyle);
				cell = row.getCell(16);
				cell.setCellValue(tecnico.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(17);
				cell.setCellValue(observacao.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);
				cell = row.getCell(18);
				cell.setCellValue(status.get(j));
				cell.setCellType(CellType.STRING);
				cell.setCellStyle(textStyle);

				j++;
			}
		} catch (NumberFormatException | IndexOutOfBoundsException e) {
			System.out.println("Escrevendo na planilha Exception -> NumberFormat or IndexOutOfBounds" + e);
		} catch (NullPointerException e) {
			System.out.println("ERRRROOO" + e);
		}
	}

	public double points(double cm) {
		return cm * 28.35;
	}

	public long twentiethsOfPoint(double cm) {
		return (long) (cm * 567);
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

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(new FlatIntelliJLaf());
					frame = new Principal();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}
}