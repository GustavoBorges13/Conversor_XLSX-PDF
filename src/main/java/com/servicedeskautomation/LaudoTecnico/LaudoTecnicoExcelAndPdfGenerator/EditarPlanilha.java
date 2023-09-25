package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.awt.image.BufferedImageOp;
import java.awt.image.ConvolveOp;
import java.awt.image.Kernel;

import javax.swing.DefaultComboBoxModel;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JSpinner;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.text.Document;
import javax.swing.text.JTextComponent;
import javax.swing.undo.UndoManager;
import com.formdev.flatlaf.FlatIntelliJLaf;
import java.awt.event.ItemListener;
import java.awt.event.ItemEvent;
import javax.swing.SwingConstants;
import javax.swing.JSeparator;

@SuppressWarnings("rawtypes")
public class EditarPlanilha extends JDialog {

	private static final long serialVersionUID = -7565309808695712383L;
	private JPanel contentPane;
	static JTextField txtLaudo;
	static JTextField txtNomeSolicitante;
	static JTextField txtUsuario;
	static JTextField txtCentroDeCusto;
	static JTextField txtItem;
	static JTextField txtAtivo;
	static JTextField txtHostname;
	static JTextField txtModelo;
	static JTextField txtServiceTag;
	static JTextField txtDdmmyyyy;
	static JTextField txtCpu;
	static JTextField txtNomeDoTecnico;
	static JTextField txtObservao;
	static JTextField txtStatus;
	static JComboBox comboBoxFabricante;
	static JComboBox comboBoxStorage;
	static JComboBox comboBoxQuantidade;
	static JSpinner spinner_memoria;
	private JLabel lblFabricante;
	private JLabel lblDataAquisicao;
	private JLabel lblsStorage;
	private JLabel lblObservacao;
	private JLabel lblStatus;
	static Component c; // contro
	static JComboBox comboBoxDispositivo;
	private int linhaSelecionada;
	private JButton btnSave;
	static boolean finalizado;
	private static int tempIndex;
	private Image img_help = new ImageIcon(SplashAnimation.class.getResource("/resources/help-icon.png")).getImage()
			.getScaledInstance(20, 20, Image.SCALE_SMOOTH);
	private BufferedImage convolvedImage;
	private JLabel lblHelp;
	private JSeparator separator_1;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(new FlatIntelliJLaf());
					EditarPlanilha frame = new EditarPlanilha();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	@SuppressWarnings({ "unchecked" })
	public EditarPlanilha() {
		setModalExclusionType(ModalExclusionType.APPLICATION_EXCLUDE);
		finalizado = false;
		setTitle("Editar informações da planilha");
		setResizable(false);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 515, 643);

		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		// Definindo uma posicao para a janela aparecer...
		/*
		 * GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
		 * GraphicsDevice defaultScreen = ge.getDefaultScreenDevice();
		 * 
		 * @SuppressWarnings("unused") Rectangle rect =
		 * defaultScreen.getDefaultConfiguration().getBounds(); // rect.getMaxX() -
		 * getWidth();
		 * 
		 * int tolerancia = -150; int x = 0; int y = ((getWidth()) / 2) + tolerancia;
		 * setLocation(x, y);
		 */
		setLocationRelativeTo(null);
		setVisible(true);

		setContentPane(contentPane);
		contentPane.setLayout(null);

		lblHelp = new JLabel("Atalhos");
		lblHelp.setHorizontalTextPosition(SwingConstants.LEFT);
		lblHelp.setHorizontalAlignment(SwingConstants.RIGHT);
		lblHelp.setIcon(new ImageIcon(img_help));
		lblHelp.setBounds(382, 0, 107, 26);
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
		// Crie uma instância personalizada de TitledBorder com fonte normal
        TitledBorder titledBorder = new TitledBorder(
            new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
            ": : Modo edi\u00E7\u00E3o : :",
            TitledBorder.LEADING,
            TitledBorder.TOP,
            new Font("Dialog", Font.PLAIN, 12), // Defina a fonte com estilo normal
            new Color(0, 0, 0)
        );
		panel.setBorder(titledBorder);
		panel.setBounds(10, 22, 479, 571);
		contentPane.add(panel);
		panel.setLayout(null);

		txtLaudo = new JTextField();
		txtLaudo.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtLaudo.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtLaudo.setForeground(Color.BLACK);
		txtLaudo.setBounds(15, 49, 126, 29);
		panel.add(txtLaudo);
		txtLaudo.setColumns(10);

		txtNomeSolicitante = new JTextField();
		txtNomeSolicitante.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtNomeSolicitante.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtNomeSolicitante.setForeground(Color.BLACK);
		txtNomeSolicitante.setColumns(10);
		txtNomeSolicitante.setBounds(151, 49, 318, 29);
		panel.add(txtNomeSolicitante);

		txtUsuario = new JTextField();
		txtUsuario.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtUsuario.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtUsuario.setForeground(Color.BLACK);
		txtUsuario.setColumns(10);
		txtUsuario.setBounds(15, 102, 126, 29);
		panel.add(txtUsuario);

		txtCentroDeCusto = new JTextField();
		txtCentroDeCusto.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtCentroDeCusto.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtCentroDeCusto.setForeground(Color.BLACK);
		txtCentroDeCusto.setColumns(10);
		txtCentroDeCusto.setBounds(150, 102, 319, 29);
		panel.add(txtCentroDeCusto);

		txtItem = new JTextField();
		txtItem.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtItem.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtItem.setForeground(Color.BLACK);
		txtItem.setColumns(10);
		txtItem.setBounds(15, 155, 339, 29);
		panel.add(txtItem);

		comboBoxQuantidade = new JComboBox();
		comboBoxQuantidade.setFont(new Font("Dialog", Font.PLAIN, 12));
		comboBoxQuantidade.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (comboBoxQuantidade.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		comboBoxQuantidade.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				requestFocus();
			}
		});
		comboBoxQuantidade.setForeground(Color.RED);
		comboBoxQuantidade.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (comboBoxQuantidade.getSelectedIndex() == 0) {
					comboBoxQuantidade.setForeground(Color.BLACK);
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (comboBoxQuantidade.getSelectedIndex() == 0) {
					comboBoxQuantidade.setForeground(Color.RED);
				} else {
					comboBoxQuantidade.setForeground(Color.BLACK);
				}
			}
		});
		comboBoxQuantidade.setMaximumRowCount(6);
		comboBoxQuantidade.setModel(new DefaultComboBoxModel(
				new String[] { "Selecionar", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" }));
		comboBoxQuantidade.setBounds(364, 155, 105, 29);
		panel.add(comboBoxQuantidade);

		JLabel lblQTD = new JLabel("Quantidade *");
		lblQTD.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblQTD.setBounds(365, 131, 104, 23);
		panel.add(lblQTD);

		txtAtivo = new JTextField();
		txtAtivo.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtAtivo.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtAtivo.setForeground(Color.BLACK);
		txtAtivo.setColumns(10);
		txtAtivo.setBounds(15, 207, 114, 29);
		panel.add(txtAtivo);

		txtHostname = new JTextField();
		txtHostname.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtHostname.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtHostname.setForeground(Color.BLACK);
		txtHostname.setColumns(10);
		txtHostname.setBounds(244, 207, 110, 29);
		panel.add(txtHostname);

		comboBoxFabricante = new JComboBox();
		comboBoxFabricante.setFont(new Font("Dialog", Font.PLAIN, 12));
		comboBoxFabricante.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (comboBoxFabricante.getSelectedIndex() == 0) {
					comboBoxFabricante.setForeground(Color.RED);
				} else {
					comboBoxFabricante.setForeground(Color.BLACK);
				}
			}
		});
		comboBoxFabricante.setForeground(Color.RED);
		comboBoxFabricante.setEditable(true);
		comboBoxFabricante.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (comboBoxFabricante.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		comboBoxFabricante.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				requestFocus();
			}
		});
		comboBoxFabricante.setModel(new DefaultComboBoxModel(new String[] { "Selecionar", "Dell Inc", "Lenovo" }));
		comboBoxFabricante.setMaximumRowCount(3);
		comboBoxFabricante.setBounds(364, 207, 105, 29);
		panel.add(comboBoxFabricante);

		lblFabricante = new JLabel("Fabricante *");
		lblFabricante.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblFabricante.setBounds(364, 184, 105, 29);
		panel.add(lblFabricante);

		txtModelo = new JTextField();
		txtModelo.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtModelo.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtModelo.setForeground(Color.BLACK);
		txtModelo.setColumns(10);
		txtModelo.setBounds(15, 264, 152, 29);
		panel.add(txtModelo);

		txtServiceTag = new JTextField();
		txtServiceTag.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtServiceTag.setToolTipText("Caso não tenha, coloque: N/A ou N/D");
		txtServiceTag.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtServiceTag.setForeground(Color.BLACK);
		txtServiceTag.setColumns(10);
		txtServiceTag.setBounds(177, 264, 129, 29);
		panel.add(txtServiceTag);

		txtDdmmyyyy = new JTextField();
		txtDdmmyyyy.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtDdmmyyyy.setToolTipText("Caso não tenha, coloque: N/A ou N/D");
		txtDdmmyyyy.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtDdmmyyyy.setForeground(new Color(255, 114, 118));
		txtDdmmyyyy.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtDdmmyyyy.getText().equals("dd/MM/yyyy")) {
					txtDdmmyyyy.setText("");
					txtDdmmyyyy.setForeground(Color.BLACK);
				} else {
					txtDdmmyyyy.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtDdmmyyyy.getText().equals("")) {
					txtDdmmyyyy.setForeground(new Color(255, 114, 118));
					txtDdmmyyyy.setText("dd/MM/yyyy");
				}

			}
		});
		txtDdmmyyyy.setText("dd/MM/yyyy");
		txtDdmmyyyy.setColumns(10);
		txtDdmmyyyy.setBounds(314, 264, 155, 29);
		panel.add(txtDdmmyyyy);

		lblDataAquisicao = new JLabel("Data aquisição *");
		lblDataAquisicao.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblDataAquisicao.setBounds(314, 236, 155, 29);
		panel.add(lblDataAquisicao);

		txtCpu = new JTextField();
		txtCpu.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtCpu.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtCpu.setForeground(Color.BLACK);
		txtCpu.setColumns(10);
		txtCpu.setBounds(15, 317, 454, 29);
		panel.add(txtCpu);

		lblsStorage = new JLabel("Storage (GB) *");
		lblsStorage.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblsStorage.setBounds(15, 346, 132, 29);
		panel.add(lblsStorage);

		JLabel lblMemoria = new JLabel("Memória (GB) *");
		lblMemoria.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblMemoria.setBounds(157, 346, 115, 29);
		panel.add(lblMemoria);

		spinner_memoria = new JSpinner();
		spinner_memoria.setFont(new Font("Dialog", Font.PLAIN, 12));
		spinner_memoria.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		spinner_memoria
				.setModel(new SpinnerNumberModel(Integer.valueOf(0), Integer.valueOf(0), null, Integer.valueOf(1)));
		c = spinner_memoria.getEditor().getComponent(0);
		c.setForeground(Color.RED);
		spinner_memoria.addChangeListener(new ChangeListener() {
			public void stateChanged(ChangeEvent arg0) {

				if (((int) spinner_memoria.getValue()) > 0) {
					c.setForeground(Color.BLACK);
				} else {

					c.setForeground(Color.RED);

				}
			}
		});

		comboBoxStorage = new JComboBox();
		comboBoxStorage.setFont(new Font("Dialog", Font.PLAIN, 12));
		comboBoxStorage.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (comboBoxStorage.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		comboBoxStorage.setMaximumRowCount(9);
		comboBoxStorage.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				requestFocus();
			}
		});
		comboBoxStorage.setForeground(Color.RED);
		comboBoxStorage.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (comboBoxStorage.getSelectedIndex() == 0) {
					comboBoxStorage.setForeground(Color.BLACK);
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (comboBoxStorage.getSelectedIndex() == 0) {
					comboBoxStorage.setForeground(Color.RED);
				} else {
					comboBoxStorage.setForeground(Color.BLACK);
				}
			}
		});
		comboBoxStorage.setModel(new DefaultComboBoxModel(new String[] { "Selecionar", "180 HD", "250 HD", "300 HD",
				"500 HD", "750 HD", "1000 HD", "120 SSD", "240 SSD", "256 SSD-NVMe" }));
		comboBoxStorage.setBounds(15, 371, 132, 29);
		panel.add(comboBoxStorage);

		spinner_memoria.setBounds(157, 371, 115, 29);
		panel.add(spinner_memoria);

		txtNomeDoTecnico = new JTextField();
		txtNomeDoTecnico.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtNomeDoTecnico.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtNomeDoTecnico.setForeground(Color.BLACK);
		txtNomeDoTecnico.setColumns(10);
		txtNomeDoTecnico.setBounds(15, 435, 454, 29);
		panel.add(txtNomeDoTecnico);

		txtObservao = new JTextField();
		txtObservao.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtObservao.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtObservao.setForeground(Color.BLACK);
		txtObservao.setColumns(10);
		txtObservao.setBounds(15, 486, 197, 29);
		panel.add(txtObservao);

		txtStatus = new JTextField();
		txtStatus.setFont(new Font("Dialog", Font.PLAIN, 12));
		txtStatus.addKeyListener(new KeyAdapter() {
			@Override
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});
		txtStatus.setForeground(Color.BLACK);
		txtStatus.setColumns(10);
		txtStatus.setBounds(220, 486, 249, 29);
		panel.add(txtStatus);

		lblObservacao = new JLabel("Observação");
		lblObservacao.setForeground(Color.BLACK);
		lblObservacao.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblObservacao.setBounds(15, 462, 197, 29);
		panel.add(lblObservacao);

		lblStatus = new JLabel("Status");
		lblStatus.setForeground(Color.BLACK);
		lblStatus.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblStatus.setBounds(220, 462, 249, 29);
		panel.add(lblStatus);

		// Botao salvar
		btnSave = new JButton("Salvar");
		btnSave.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (btnSave.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		btnSave.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				linhaSelecionada = Principal.table.getSelectedRow();
				boolean flag = false;
				String sLaudo = txtLaudo.getText();
				// String sQtd = comboBoxQuantidade.getSelectedIndex()+"";
				String Memoria = spinner_memoria.getValue().toString();
				String dateString = txtDdmmyyyy.getText();

				// Manipulando os arraylist (dadabase)
				if (txtLaudo.getText().equals("") || txtNomeSolicitante.getText().equals("")
						|| txtUsuario.getText().equals("") || txtCentroDeCusto.getText().equals("")
						|| txtItem.getText().equals("") || comboBoxQuantidade.getSelectedIndex() == 0
						|| txtAtivo.getText().equals("") || comboBoxDispositivo.getSelectedIndex() == 0
						|| txtHostname.getText().equals("") || comboBoxFabricante.getSelectedIndex() == 0
						|| txtModelo.getText().equals("") || txtServiceTag.getText().equals("")
						|| txtDdmmyyyy.getText().equals("dd/mm/yyyy") || txtCpu.getText().equals("")
						|| comboBoxStorage.getSelectedIndex() == 0 || ((int) spinner_memoria.getValue()) == 0
						|| txtNomeDoTecnico.getText().equals("")) {

					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"Existem alguns campos obrigatórios (*) em brancos ou você\nnão selecionou nenhum item das caixas de seleções.\nPor favor preencha os campos.",
							"Erro ao salvar...", JOptionPane.INFORMATION_MESSAGE);
					return;
				} else if (sLaudo.matches("[a-zA-Z]+")) { // Verificar se é uma string
					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"O campo de numero do laudo contem letras, por favor digite apenas numeros.",
							"Erro InputMismatchException", JOptionPane.INFORMATION_MESSAGE);
					return;
				} else if (!sLaudo.matches("[0-9]{6}")) { // Verificar se tem 6 digitos
					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"O campo de numero do laudo contem menos ou mais de 6 digitos, por favor digite 6 digitos!",
							"Erro length", JOptionPane.INFORMATION_MESSAGE);
					return;
				} else if (Memoria.matches("[a-zA-Z]+")) { // Verificar se é uma string
					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"O campo de memoria contem letras, por favor digite apenas numeros.",
							"Erro InputMismatchException", JOptionPane.INFORMATION_MESSAGE);
					return;
				} else if (!dateString.matches("\\d{2}/\\d{2}/\\d{4}") && !txtDdmmyyyy.getText().equals("N/D")
						&& !txtDdmmyyyy.getText().equals("N/A")) { // -> dd/MM/yyyy
					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"O campo de data não está no formato solicitado (dd/MM/yyyy), por favor reescreva. Ou então insira N/A ou N/D caso não tenha.",
							"Erro InputMismatchException", JOptionPane.INFORMATION_MESSAGE);
					return;
				} else {

					Principal.btnSalvarAlteracoes.setEnabled(true);
					flag = true;
					for (int coluna = 0; coluna < Principal.colunas.size(); coluna++) {

						// Manipulando os arraylist (dadabase)
						Principal.laudo.set(linhaSelecionada, txtLaudo.getText());
						Principal.nomeSolicitante.set(linhaSelecionada, txtNomeSolicitante.getText());
						Principal.usuario.set(linhaSelecionada, txtUsuario.getText());
						Principal.centroCusto.set(linhaSelecionada, txtCentroDeCusto.getText());
						Principal.item.set(linhaSelecionada, txtItem.getText());
						Principal.qtd.set(linhaSelecionada, comboBoxQuantidade.getSelectedItem().toString());
						Principal.ativo.set(linhaSelecionada, txtAtivo.getText());
						Principal.dispositivo.set(linhaSelecionada, comboBoxDispositivo.getSelectedItem().toString());
						Principal.hostname.set(linhaSelecionada, txtHostname.getText());
						Principal.fabricante.set(linhaSelecionada, comboBoxFabricante.getSelectedItem().toString());
						Principal.modelo.set(linhaSelecionada, txtModelo.getText());
						Principal.serviceTag.set(linhaSelecionada, txtServiceTag.getText());
						Principal.dataAquisicao.set(linhaSelecionada, txtDdmmyyyy.getText());
						Principal.cpu.set(linhaSelecionada, txtCpu.getText());
						Principal.storage.set(linhaSelecionada, comboBoxStorage.getSelectedItem().toString());
						Principal.memoria.set(linhaSelecionada, spinner_memoria.getValue().toString());
						Principal.tecnico.set(linhaSelecionada, txtNomeDoTecnico.getText());
						Principal.observacao.set(linhaSelecionada, txtObservao.getText());
						Principal.status.set(linhaSelecionada, txtStatus.getText());

						// Modificando os valores da tabela de forma temporaria (sem salvar na planilha
						// apenas visual)
						Principal.table.setValueAt(txtLaudo.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtNomeSolicitante.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtUsuario.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtCentroDeCusto.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtItem.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(comboBoxFabricante.getSelectedItem().toString(), linhaSelecionada,
								coluna); // BUG
						Principal.table.setValueAt(txtAtivo.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(comboBoxDispositivo.getSelectedItem().toString(), linhaSelecionada,
								coluna);
						Principal.table.setValueAt(txtHostname.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(comboBoxFabricante.getSelectedItem().toString(), linhaSelecionada,
								coluna); // BUG
						Principal.table.setValueAt(txtModelo.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtServiceTag.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtDdmmyyyy.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtCpu.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(comboBoxStorage.getSelectedItem().toString(), linhaSelecionada,
								coluna); // BUG
						Principal.table.setValueAt(spinner_memoria.getValue(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtNomeDoTecnico.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtObservao.getText(), linhaSelecionada, coluna);
						Principal.table.setValueAt(txtStatus.getText(), linhaSelecionada, coluna);
					}
				}

				if (flag) {
					// Fechamento da janela
					finalizado = true;
					Principal.flagSaved = false;
					
					// volta a janela principal para o estado anterior
					dispose();
				}
			}
		});
		btnSave.setBounds(177, 537, 129, 23);
		panel.add(btnSave);

		JLabel lblChamado = new JLabel("Chamado *");
		lblChamado.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblChamado.setBounds(15, 27, 126, 20);
		panel.add(lblChamado);

		JLabel lblNomeDoColaborador = new JLabel("Nome do colaborador *");
		lblNomeDoColaborador.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblNomeDoColaborador.setBounds(151, 27, 318, 20);
		panel.add(lblNomeDoColaborador);

		JLabel lblUsurioDeRede = new JLabel("Usuário de rede *");
		lblUsurioDeRede.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblUsurioDeRede.setBounds(15, 78, 126, 25);
		panel.add(lblUsurioDeRede);

		JLabel lblDescrioDoItem = new JLabel("Descrição do item *");
		lblDescrioDoItem.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblDescrioDoItem.setBounds(15, 132, 339, 21);
		panel.add(lblDescrioDoItem);

		JLabel lblCentroDeCusto = new JLabel("Centro de custo *");
		lblCentroDeCusto.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblCentroDeCusto.setBounds(151, 78, 318, 25);
		panel.add(lblCentroDeCusto);

		JLabel lblAtivo = new JLabel("Ativo *");
		lblAtivo.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblAtivo.setBounds(15, 184, 114, 29);
		panel.add(lblAtivo);

		JLabel lblDispositivo = new JLabel("Dispositivo *");
		lblDispositivo.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblDispositivo.setBounds(135, 184, 96, 29);
		panel.add(lblDispositivo);

		JLabel lblHostname = new JLabel("Hostname *");
		lblHostname.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblHostname.setBounds(244, 184, 110, 29);
		panel.add(lblHostname);

		JLabel lblModelo = new JLabel("Modelo *");
		lblModelo.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblModelo.setBounds(15, 236, 152, 29);
		panel.add(lblModelo);

		JLabel lblServiceTag = new JLabel("Service TAG *");
		lblServiceTag.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblServiceTag.setBounds(177, 237, 129, 29);
		panel.add(lblServiceTag);

		JLabel lblEspecificaesDoProcessador = new JLabel("Especificações do processador *");
		lblEspecificaesDoProcessador.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblEspecificaesDoProcessador.setBounds(15, 294, 454, 23);
		panel.add(lblEspecificaesDoProcessador);

		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico (elaborador do laudo) *");
		lblNomeDoTcnico.setFont(new Font("Dialog", Font.PLAIN, 12));
		lblNomeDoTcnico.setBounds(15, 411, 454, 29);
		panel.add(lblNomeDoTcnico);

		comboBoxDispositivo = new JComboBox();
		comboBoxDispositivo.setFont(new Font("Dialog", Font.PLAIN, 12));
		comboBoxDispositivo.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent e) {
				if (comboBoxDispositivo.isEnabled())
					setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
			}
		});
		comboBoxDispositivo.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				requestFocus();
			}
		});
		comboBoxDispositivo.setForeground(Color.RED);
		comboBoxDispositivo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (comboBoxDispositivo.getSelectedIndex() == 0) {
					comboBoxDispositivo.setForeground(Color.BLACK);
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (comboBoxDispositivo.getSelectedIndex() == 0) {
					comboBoxDispositivo.setForeground(Color.RED);
				} else {
					comboBoxDispositivo.setForeground(Color.BLACK);
				}
			}
		});
		comboBoxDispositivo.setModel(new DefaultComboBoxModel(new String[] { "Selecionar", "Desktop", "Notebook" }));
		comboBoxDispositivo.setMaximumRowCount(3);
		comboBoxDispositivo.setForeground(Color.RED);
		comboBoxDispositivo.setBounds(138, 207, 96, 29);
		panel.add(comboBoxDispositivo);
		
		JSeparator separator = new JSeparator();
		separator.setBounds(15, 525, 454, 1);
		panel.add(separator);
		
		separator_1 = new JSeparator();
		separator_1.setBounds(15, 411, 454, 1);
		panel.add(separator_1);
		requestFocus();

		addKeyListener(new KeyHandler() {
			public void keyPressed(KeyEvent e) {
				if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
					dispose();
				}
				if (e.getKeyCode() == KeyEvent.VK_ENTER) {
					btnSave.doClick();
				}
			}
		});

		// Adiciona a funcionalidade de desfazer/refazer em todos os componentes
		addUndoRedoFunctionality(getContentPane());
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
			} else if (component instanceof JComboBox) {
				JComboBox comboBox = (JComboBox) component;
				UndoManager undoManager = new UndoManager();
				Document docTemp = ((JTextComponent) comboBox.getEditor().getEditorComponent()).getDocument();

				docTemp.addUndoableEditListener(undoManager);

				comboBox.addKeyListener(new KeyAdapter() {
					public void keyPressed(KeyEvent e) {
						if (e.getKeyCode() == KeyEvent.VK_Z && e.isControlDown()) {
							tempIndex = comboBox.getSelectedIndex();
							comboBox.setSelectedIndex(0);
						} else if (e.getKeyCode() == KeyEvent.VK_Y && e.isControlDown()) {
							comboBox.setSelectedIndex(tempIndex);
						}
					}
				});
			} else if (component instanceof JSpinner) {
				JSpinner spinner = (JSpinner) component;
				UndoManager undoManager = new UndoManager();
				// Obtém a interface gráfica do editor do JSpinner
				JTextComponent editorComponent = ((JSpinner.DefaultEditor) spinner.getEditor()).getTextField();
				Document docTemp = editorComponent.getDocument();
				docTemp.addUndoableEditListener(undoManager);

				editorComponent.addKeyListener(new KeyAdapter() {
					public void keyPressed(KeyEvent e) {
						if (e.getKeyCode() == KeyEvent.VK_Z && e.isControlDown()) {
							if (undoManager.canUndo()) {
								undoManager.undo();
								undoManager.undo();
							}
						} else if (e.getKeyCode() == KeyEvent.VK_Y && e.isControlDown()) {
							if (undoManager.canRedo()) {
								undoManager.redo();
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
