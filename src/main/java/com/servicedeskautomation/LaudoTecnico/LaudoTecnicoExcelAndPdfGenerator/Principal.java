package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
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
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JInternalFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSeparator;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.KeyStroke;
import javax.swing.ListSelectionModel;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.event.InternalFrameAdapter;
import javax.swing.event.InternalFrameEvent;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.plaf.basic.BasicInternalFrameUI;
import javax.swing.table.TableCellRenderer;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import org.apache.poi.ss.usermodel.BorderStyle;
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
import java.awt.event.KeyEvent;
import java.awt.event.InputEvent;

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
	private JInternalFrame internalFrame;
	static JButton btnSalvarAlteracoes;
	private JButton btnGerarArquivoPdf;
	static JButton btnEditar;
	private JTextPane txtpnAplicaoDesenvolvidaPara;;
	public int KEYCODE_ESC = 27;

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
	static Principal frame;
	private JButton btnRemover;
	private int qtdTemporaria;

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
		table.getColumnModel().getColumn(16).setPreferredWidth(110); // coluna TECNICO
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
		table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

		// Realiza a configuracao de fontes
		table.getTableHeader().setFont(new Font("Dialog", Font.BOLD | Font.ITALIC, 10));
		table.setFont(new Font("Dialog", Font.PLAIN, 10));
	}

	public Principal() {
		setIconImage(
				Toolkit.getDefaultToolkit().getImage(Principal.class.getResource("/resource/icon-ServiceDesk.png")));
		setAutoRequestFocus(false);
		setTitle("E-ServiceDesk application");
		setResizable(false);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 755, 580);
		setLocationRelativeTo(null);

		JMenuBar menuBar = new JMenuBar();
		setJMenuBar(menuBar);

		JMenu mnNewMenu = new JMenu("Ajuda");
		mnNewMenu.setMnemonic('A');
		menuBar.add(mnNewMenu);

		JMenuItem mntmNewMenuItem = new JMenuItem("Sobre...");
		mntmNewMenuItem.setMnemonic('S');
		mntmNewMenuItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_DOWN_MASK));
		mntmNewMenuItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				internalFrame.setFocusable(true);
				internalFrame.setVisible(true);
				internalFrame.requestFocus();
				txtpnAplicaoDesenvolvidaPara.setCaretPosition(0);// Sobe para cima a barra de rolagem vertical\
				jlocal.setOpaque(false);
			}
		});

		JMenuItem mntmNewMenuItem_1 = new JMenuItem("Creditos...");
		mntmNewMenuItem_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JOptionPane.showMessageDialog(null, new MessageWithLink(
						"Conversor XLSX-PDF<br/>Versão XXX <br/><br/>Gustavo Borges<br/><a href=\"https://github.com/GustavoBorges13\">https://github.com/GustavoBorges13</a>"),
						"Informações adicionais", JOptionPane.INFORMATION_MESSAGE);
			}
		});
		mntmNewMenuItem_1.setMnemonic('C');
		mntmNewMenuItem_1.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_C, InputEvent.CTRL_DOWN_MASK));
		mnNewMenu.add(mntmNewMenuItem_1);
		mnNewMenu.add(mntmNewMenuItem);

		JMenu mnNewMenu_1 = new JMenu("Tema (Beta)");
		mnNewMenu_1.setMnemonic('T');
		menuBar.add(mnNewMenu_1);

		JCheckBoxMenuItem CheckTema1 = new JCheckBoxMenuItem("Tema 1");
		mnNewMenu_1.add(CheckTema1);

		JCheckBoxMenuItem CheckTema2 = new JCheckBoxMenuItem("Tema 2");
		mnNewMenu_1.add(CheckTema2);

		contentPane = new JPanel();
		contentPane.setOpaque(false);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		internalFrame = new JInternalFrame("Sobre a aplicação");
		internalFrame.setEnabled(false);
		internalFrame.setFocusable(false);

		// Janela Ajuda-Sobre (Sobre a aplicacao)
		BasicInternalFrameUI basicInternalFrameUI = ((javax.swing.plaf.basic.BasicInternalFrameUI) internalFrame
				.getUI());
		for (MouseListener listener : basicInternalFrameUI.getNorthPane().getMouseListeners()) {
			basicInternalFrameUI.getNorthPane().removeMouseListener(listener);
		}

		internalFrame.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				btnXLS.setVisible(false);
				table.setVisible(false);
			}

			@Override
			public void focusLost(FocusEvent e) {
				btnXLS.setVisible(true);
				table.setVisible(true);
			}
		});
		internalFrame.addInternalFrameListener(new InternalFrameAdapter() {
			@Override
			public void internalFrameClosed(InternalFrameEvent arg0) {
				jlocal.setOpaque(true);
			}
		});

		contentPane.add(internalFrame);

		internalFrame.pack();
		internalFrame.getContentPane().setLayout(null);
		internalFrame.setBounds(39, 0, 664, 508);
		internalFrame.setOpaque(true);
		internalFrame.setClosable(true);
		internalFrame.setVisible(false);

		JLabel lblNewLabel = new JLabel("E-ServiceDesk Applications - About");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 15));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBounds(10, 11, 628, 19);
		internalFrame.getContentPane().add(lblNewLabel);

		JPanel panel = new JPanel();
		panel.setBackground(new Color(240, 240, 240));
		panel.setBounds(10, 41, 628, 426);
		internalFrame.getContentPane().add(panel);
		panel.setLayout(null);
		SimpleAttributeSet attr = new SimpleAttributeSet();
		attr.copyAttributes();
		StyleConstants.setAlignment(attr, StyleConstants.ALIGN_JUSTIFIED);

		JScrollPane scrollPane_1 = new JScrollPane();
		scrollPane_1.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane_1.setBounds(0, 0, 628, 426);
		panel.add(scrollPane_1);

		txtpnAplicaoDesenvolvidaPara = new JTextPane();
		txtpnAplicaoDesenvolvidaPara.setAutoscrolls(false);
		scrollPane_1.setViewportView(txtpnAplicaoDesenvolvidaPara);

		txtpnAplicaoDesenvolvidaPara.setDisabledTextColor(new Color(0, 0, 0));
		txtpnAplicaoDesenvolvidaPara.setEnabled(false);
		txtpnAplicaoDesenvolvidaPara.setEditable(false);
		txtpnAplicaoDesenvolvidaPara.setBackground(new Color(240, 240, 240));
		txtpnAplicaoDesenvolvidaPara.setParagraphAttributes(attr, true);
		txtpnAplicaoDesenvolvidaPara.setFont(new Font("Dialog", Font.PLAIN, 11));
		txtpnAplicaoDesenvolvidaPara.setText(
				"Aplicação desenvolvida para ajudar a nossa equipe service-desk.\r\n\r\nEsté é um programa que realiza um espécime de automação, para tornar o trabalho desenvolvido mais rapido. Foi desenvolvido para fins educacionais com a intenção de obter mais conhecimentos em APIs distintas que antes eu nunca tinha visto como por exemplo APACHE POI, JXL, OpenCSV e docx4j, e claro, melhorar minhas habilidades com a linguagem de programação em um ambiente profissional.\r\n\r\nO projeto foi desenvolvido utilizando Eclipse versão de 2022-09, cujo as builds foram realizadas no MAVEN para fazer clean verify, instalar pacotes, e builds para evitar problemas de ter alguma API ausente no projeto ao transitar entre as máquinas da minha casa com a da empresa, ou seja, essa transição foi feita pelo gitbash do github para salvar os commits do projeto livremente, para mais informações acesse o meu perfil no gitHub e verifique o repositorio em Ajuda-Sobre (Ctrl+S).\r\n\r\nSobre a execução do programa, se trata de uma interface gráfica dinâmica, no qual o usuário se depara com uma primeira janela para escolher a planilha em especifico que será manipulada sem precisar utilizar o excel, sendo que TODO o codigo foi feito especialmente para este tipo de planilha, levando em considerações desde das formatações e quantidade de colunas contidas nele. \r\nAo selecionar a planilha a mesma é transposta para uma Jtable afins de tornar a tabela editavel \"como se fosse um excel\". Além disso, caso o usuario selecione alguma linha da tabela, a mesma irá habilitar opções de edições e ao clicar no botão ou clicar duas vezes na linha que deseja editar, irá abrir uma janela na lateral esquerda transcrevendo os valores selecionados para a edição do mesmo. \r\nCaso o usuario esteja satisfeito, poderá selecionar a linha em especifico e utilizar a ferramenta de EXPORTAR, onde fará será realizado uma automação, transcrevendo os dados da linha selecionada para um documento WORD formatado no padrão da empresa e logo seguinte salva-lo em PDF na pasta alvo que posteriormente será aberto automaticamente para revisão do mesmo para ser enviado ao cliente sucessivamente.\r\n\r\nAtt. Gustavo Borges.");
		JSeparator separator = new JSeparator();
		separator.setForeground(new Color(105, 105, 105));
		separator.setBounds(11, 33, 627, 2);
		internalFrame.getContentPane().add(separator);

		jlocal = new JTextField();
		jlocal.setFocusable(false);
		jlocal.setEditable(false);
		jlocal.setBackground(new Color(255, 255, 255));
		jlocal.setBounds(39, 49, 521, 39);
		contentPane.add(jlocal);
		jlocal.setColumns(10);
		jlocal.setOpaque(true);
		btnXLS = new JButton("XLSX");

		// BUTTON CHOOSER FILE
		btnXLS.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fc = new JFileChooser();
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
				} else {
					pathfileExcel = fc.getSelectedFile();
					jlocal.setText(pathfileExcel.toString().trim());
					jtitulo.setText(fc.getSelectedFile().getName());
					btnPreencher.setEnabled(true); // Habilita o botao
					
					// Limpa selecoes
					table.clearSelection();
					requestFocus();
					
					// Limpa tabela se estiver aberta antes
					limpaListas();
					table.createDefaultColumnsFromModel();
				}

				// Habilita os botoes auxiliares para controlar a planilha
				btnAddLinha.setEnabled(false);
				btnEditar.setEnabled(false);
				btnRemover.setEnabled(false);
				btnSalvarAlteracoes.setEnabled(false);
				btnGerarArquivoPdf.setEnabled(false);
				btnPreencher.setEnabled(true);

				
			}
		});
		btnXLS.setBounds(583, 49, 113, 39);
		contentPane.add(btnXLS);

		// BUTTON FILL
		btnPreencher = new JButton("Preencher");
		btnPreencher.setFocusable(false);
		btnPreencher.setEnabled(false);

		// Button fill/preencher
		btnPreencher.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				btnSalvarAlteracoes.setEnabled(false);
				XSSFWorkbook work;
				XSSFSheet sheet;
				XSSFRow row;

				// Limpeza dos dados salvos nas listas é como se fosse um F5
				limpaListas();

				try {
					int i = 0; // primeira linha
					int pos = 0;

					// Abre o arquivo excel
					work = new XSSFWorkbook(new FileInputStream(pathfileExcel));

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
					JOptionPane.showMessageDialog(null, "Arquivo Excel não encontrado!\nErro: " + e1);
				} catch (NullPointerException e2) {
					System.out.println("Debug error: " + e2 + " Não há mais colunas no excel...");
				} catch (IOException e3) {
					JOptionPane.showMessageDialog(null, "Arquivo invalido!\nErro: " + e3);
				} finally {
					try {
						int i = 1; // varredura a partir da segunda linha ~ ignora cabeçalho
						int pos = 0; // valor default

						// Abre o arquivo excel
						work = new XSSFWorkbook(new FileInputStream(pathfileExcel));

						// Junta todas as planilhas deste arquivo.
						// Planilha 1 = 0 ou usar getSheet ("NOME DA Planilha")
						sheet = work.getSheetAt(0);

						// Linha de referencia (selecionar a partir de uma determinada linha...)
						row = sheet.getRow(i);
						if (colunas.size() == 19) {
							// Copia as informações das linhas da planilhas para arrayslists.
							while (((row = sheet.getRow(i)) != null && row.getCell(0).getCellType() != CellType.BLANK)
									|| ((row = sheet.getRow(i)) != null
											&& row.getCell(1).getCellType() != CellType.BLANK)
									|| ((row = sheet.getRow(i)) != null
											&& row.getCell(2).getCellType() != CellType.BLANK)) {
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

								//Tratamento de exeptions
								try {
									//Tentativa de salvar o valor da celular (DATE) na lista do tipo String
									if (DateUtil.isCellDateFormatted(row.getCell(12))) {
										Date date = row.getCell(12).getDateCellValue();
										SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
										String formattedDate = sdf2.format(date);
										dataAquisicao.add(formattedDate);
										// JOptionPane.showMessageDialog(null, "DateString -> " + date + "\nFormatted
										// Date -> " + formattedDate);
									
									//Caso o valor da celula não seja no formato DATE, verificar se é do tipo numero
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
									
								//Erro em caso de nâo ser reconhecido o tipo numerico
								} catch (IllegalStateException e2) {
									System.out.println("Debug Date Type -> " + e2);
								} finally {
									//Verifica se o tipo é string
									if (row.getCell(12).getCellType() == CellType.STRING) {
										String dateString = row.getCell(12).getStringCellValue();
										try {
											SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

											Date date = sdf.parse(dateString);
											SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
											String formattedDate = sdf2.format(date);
											dataAquisicao.add(formattedDate);
											// JOptionPane.showMessageDialog(null, "DateString -> " + dateString +
											// "\nDate
											// -> "
											// + date + "\nFormatted Date -> " + formattedDate);
										} catch (ParseException e2) {
											System.out.println("Debug date -> " + e2);
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
								observacao.add(row.getCell(17).getStringCellValue());
								status.add(row.getCell(18).getStringCellValue());

								i++;
							}
							qtdTemporaria = laudo.size();
							work.close();

						} else {
							throw new java.lang.ArrayIndexOutOfBoundsException();
						}

					} catch (NullPointerException e1) {
						System.out.println("Debug error: " + e1 + " Não há mais linhas com conteúdo no excel...");
					} catch (IOException e2) {
						JOptionPane.showMessageDialog(null, "Arquivo invalido!\nErro: " + e2);
					} catch (java.lang.ArrayIndexOutOfBoundsException e3) {
						JOptionPane.showMessageDialog(null,
								"Esta não é a planilha que utilizamos em 2022-2023.\nPor favor, abra a planilha com o modelo padrão utilizado pois,"
										+ " este codigo foi feito especificamente para \nser utilizado com esse tipo de planilha devido as formatações"
										+ " e quantidade de colunas.\nErro: " + e3);
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

				// Preenche a tabela
				preencherTabelaProprietario();

				// Habilita os botoes auxiliares para controlar a planilha
				btnAddLinha.setEnabled(true);
				btnEditar.setEnabled(false);
				btnRemover.setEnabled(false);
			}
		});

		btnPreencher.setBounds(583, 131, 113, 39);
		contentPane.add(btnPreencher);

		jtitulo = new JTextField();
		jtitulo.setDisabledTextColor(new Color(0, 0, 0));
		jtitulo.setEnabled(false);
		jtitulo.setColumns(10);
		jtitulo.setBounds(39, 131, 521, 39);
		contentPane.add(jtitulo);

		JLabel lblTituloPlanilha = new JLabel("Titulo da planilha");
		lblTituloPlanilha.setBounds(39, 106, 105, 14);
		contentPane.add(lblTituloPlanilha);

		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(39, 195, 657, 279);
		contentPane.add(scrollPane);

		table = new JTable();
		table.setAutoscrolls(false);
		table.addMouseListener(new MouseAdapter() {
			boolean isAlreadyOneClick = false;

			@Override
			public void mouseClicked(MouseEvent arg0) {
				if (isAlreadyOneClick) {
					int linhaSelecionada = table.getSelectedRow();
					EditarPlanilha frame = new EditarPlanilha();
					int toleranciaHD_SSD = 0;
					int pos = 0; // valor default

					// Foco total para a nova janela aberta
					frame.setVisible(true);
					frame.requestFocus();

					// Deixa a janela principal desativada
					Principal.this.setEnabled(false);

					// Evento pra verificar se a janela de edicao foi fechada
					frame.addWindowListener(new WindowAdapter() {
						@Override
						public void windowClosing(WindowEvent e) {
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

							// frame.setVisible(false);
						}

						@Override
						public void windowClosed(WindowEvent e) {
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

							// frame.setVisible(false);

						}

					});

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
							if(EditarPlanilha.comboBoxFabricante.getItemAt(i).toString().toUpperCase().contains(fabricante.get(linhaSelecionada).toUpperCase())) {
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
				if(table.isRowSelected(linhaSelecionada)) {
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
		btnAddLinha.setEnabled(false);
		btnAddLinha.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				// Habilita botoes
				btnEditar.setEnabled(true);
				btnRemover.setEnabled(true);
				btnSalvarAlteracoes.setEnabled(true);

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
		btnEditar.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
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
		btnEditar.setEnabled(false);
		btnEditar.setBounds(39, 485, 105, 23);
		contentPane.add(btnEditar);

		// BUTTON SAVE
		btnSalvarAlteracoes = new JButton("Salvar alterações");
		btnSalvarAlteracoes.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				XSSFWorkbook work = null;
				int reply = JOptionPane.showConfirmDialog(null, "Você realmente deseja salvar (S/N)?", "Salvamento",
						JOptionPane.YES_NO_OPTION);
				if (reply == JOptionPane.YES_OPTION) {
					try {
						// Abra o arquivo xlsx existente
						int g = 1;
						work = new XSSFWorkbook(new FileInputStream(pathfileExcel));

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
						if (laudo.size() <= qtdTemporaria)
							sheet.copyRows(sheet.getLastRowNum() - qtdTemporaria, sheet.getLastRowNum(), 1,
									new CellCopyPolicy());
						else
							sheet.copyRows(sheet.getLastRowNum() - laudo.size(), sheet.getLastRowNum(), 1,
									new CellCopyPolicy());

						// Calcula g
						while (((row = sheet.getRow(g)) != null && row.getCell(3).getCellType() == CellType.BLANK)) {
							g++;
						}

						// SALVANDO NEW DADOS NA PLANILHA!
						escreverPlanilha(work, cell, sheet, textStyle);

						// Salvar alterações na planilha!
						work.write(new FileOutputStream(pathfileExcel));

						// Adicionando mais linhas se estiver perto da linha de aviso (AMARELO)
						int limitador = 3;
						if ((g - laudo.size()) <= limitador) {
							for (int i = 0; i <= limitador; i++) {
								addRow(work, sheet, 1, laudo.size() + 1, textStyle);
							}
						}

						try {
							// Salvar alterações na planilha!
							work.write(new FileOutputStream(pathfileExcel));
							work.close(); // Fecha a planilha
							JOptionPane.showMessageDialog(null, "As alterações foram salvas com suceeso!");

							// Desabilita botoes
							btnSalvarAlteracoes.setEnabled(false);
							btnEditar.setEnabled(false);
							btnGerarArquivoPdf.setEnabled(false);

						} catch (FileNotFoundException e) {
							JOptionPane.showMessageDialog(null, "Erro ao salvar a planilha!\n" + e);
						}
					} catch (IOException e) {
						e.printStackTrace();
					}
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
		btnGerarArquivoPdf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				pathfileWord = new File("Data");
				pathfileWord.mkdirs();
				File[] listOfFiles = pathfileWord.listFiles();
				File file = null;
				int numberOfFiles = listOfFiles.length;
				/*
				 * // AUTO GENERATED FILES // Cria um especime de backup na pasta data para
				 * salvar os arquivos docx if (numberOfFiles >= 0) { file = new
				 * File(pathfileWord, "ModeloLaudo-AutoGenerated-" + numberOfFiles + ".docx"); }
				 */

				
			}
		});
		btnGerarArquivoPdf.setEnabled(false);
		btnGerarArquivoPdf.setBounds(532, 485, 164, 23);
		contentPane.add(btnGerarArquivoPdf);

		// BUTTON REMOVE
		btnRemover = new JButton("Remover");
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
				}
			}
		});
		btnRemover.setEnabled(false);
		btnRemover.setBounds(269, 485, 105, 23);
		contentPane.add(btnRemover);

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

	public static void limpaListas() {
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
		dados.remove(dados.size() - 1);
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

	private static void addRow(XSSFWorkbook workbook, XSSFSheet worksheet, int sourceRowNum, int destinationRowNum,
			XSSFCellStyle textStyle) {
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

	public static void escreverPlanilha(XSSFWorkbook work, XSSFCell cell, XSSFSheet sheet, XSSFCellStyle textStyle) {
		int j = 0;
		try {
			// Transcreve os valores do arraylist para as celulas da planilha...
			for (int i = 1; i <= laudo.size(); i++) { // Obtendo a linha desejada
				XSSFRow row = sheet.getRow(i);

				// SALVANDO NEW DADOS NA PLANILHA!
				// Alterando valores das celulas...
				cell = row.getCell(0); // Obtem a celula desejada -> Coluna
				cell.setCellValue(Integer.parseInt(laudo.get(j))); // Alterando o valor das células da
																	// planilha
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
		}
	}

	public double points(double cm) {
		return cm * 28.35;
	}

	public long twentiethsOfPoint(double cm) {
		return (long) (cm * 567);
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
