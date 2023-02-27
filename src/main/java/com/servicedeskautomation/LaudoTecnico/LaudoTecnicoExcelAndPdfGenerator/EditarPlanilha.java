package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import javax.swing.DefaultComboBoxModel;
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
		setBounds(100, 100, 465, 600);

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

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Modo edi\u00E7\u00E3o : :", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(10, 11, 429, 539);
		contentPane.add(panel);
		panel.setLayout(null);

		txtLaudo = new JTextField();
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
		txtLaudo.setBounds(15, 37, 99, 29);
		panel.add(txtLaudo);
		txtLaudo.setColumns(10);

		txtNomeSolicitante = new JTextField();
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
		txtNomeSolicitante.setBounds(124, 37, 288, 29);
		panel.add(txtNomeSolicitante);

		txtUsuario = new JTextField();
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
		txtUsuario.setBounds(15, 90, 98, 29);
		panel.add(txtUsuario);

		txtCentroDeCusto = new JTextField();
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
		txtCentroDeCusto.setBounds(123, 90, 289, 29);
		panel.add(txtCentroDeCusto);

		txtItem = new JTextField();
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
		txtItem.setBounds(15, 144, 296, 29);
		panel.add(txtItem);

		comboBoxQuantidade = new JComboBox();
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
		comboBoxQuantidade.setBounds(316, 144, 96, 29);
		panel.add(comboBoxQuantidade);

		JLabel lblQTD = new JLabel("Quantidade *");
		lblQTD.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblQTD.setBounds(316, 119, 96, 29);
		panel.add(lblQTD);

		txtAtivo = new JTextField();
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
		txtAtivo.setBounds(15, 198, 99, 29);
		panel.add(txtAtivo);

		txtHostname = new JTextField();
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
		txtHostname.setBounds(220, 198, 91, 29);
		panel.add(txtHostname);

		comboBoxFabricante = new JComboBox();
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
		comboBoxFabricante.setForeground(Color.RED);
		comboBoxFabricante.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (comboBoxFabricante.getSelectedIndex() == 0) {
					comboBoxFabricante.setForeground(Color.BLACK);
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (comboBoxFabricante.getSelectedIndex() == 0) {
					comboBoxFabricante.setForeground(Color.RED);
				} else {
					comboBoxFabricante.setForeground(Color.BLACK);
				}
			}
		});
		comboBoxFabricante.setModel(new DefaultComboBoxModel(new String[] { "Selecionar", "Dell Inc", "Lenovo" }));
		comboBoxFabricante.setMaximumRowCount(3);
		comboBoxFabricante.setBounds(316, 198, 96, 29);
		panel.add(comboBoxFabricante);

		lblFabricante = new JLabel("Fabricante *");
		lblFabricante.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblFabricante.setBounds(316, 175, 96, 29);
		panel.add(lblFabricante);

		txtModelo = new JTextField();
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
		txtModelo.setBounds(15, 258, 152, 29);
		panel.add(txtModelo);

		txtServiceTag = new JTextField();
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
		txtServiceTag.setBounds(177, 258, 129, 29);
		panel.add(txtServiceTag);

		txtDdmmyyyy = new JTextField();
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
		txtDdmmyyyy.setBounds(314, 258, 98, 29);
		panel.add(txtDdmmyyyy);

		lblDataAquisicao = new JLabel("Data aquisição *");
		lblDataAquisicao.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDataAquisicao.setBounds(314, 236, 98, 29);
		panel.add(lblDataAquisicao);

		txtCpu = new JTextField();
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
		txtCpu.setBounds(15, 309, 397, 29);
		panel.add(txtCpu);

		lblsStorage = new JLabel("Storage (GB) *");
		lblsStorage.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblsStorage.setBounds(15, 337, 132, 29);
		panel.add(lblsStorage);

		JLabel lblMemoria = new JLabel("Memória (GB) *");
		lblMemoria.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblMemoria.setBounds(157, 337, 79, 29);
		panel.add(lblMemoria);

		spinner_memoria = new JSpinner();
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
		comboBoxStorage.setBounds(15, 362, 132, 29);
		panel.add(comboBoxStorage);

		spinner_memoria.setBounds(157, 362, 99, 29);
		panel.add(spinner_memoria);

		txtNomeDoTecnico = new JTextField();
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
		txtNomeDoTecnico.setBounds(15, 415, 397, 29);
		panel.add(txtNomeDoTecnico);

		txtObservao = new JTextField();
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
		txtObservao.setBounds(15, 466, 190, 29);
		panel.add(txtObservao);

		txtStatus = new JTextField();
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
		txtStatus.setBounds(220, 466, 192, 29);
		panel.add(txtStatus);

		lblObservacao = new JLabel("Observação");
		lblObservacao.setForeground(Color.BLACK);
		lblObservacao.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblObservacao.setBounds(15, 442, 85, 29);
		panel.add(lblObservacao);

		lblStatus = new JLabel("Status");
		lblStatus.setForeground(Color.BLACK);
		lblStatus.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblStatus.setBounds(220, 442, 83, 29);
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
				} else if (!dateString.matches("\\d{2}/\\d{2}/\\d{4}")) { // -> dd/MM/yyyy
					JOptionPane.showMessageDialog(EditarPlanilha.this,
							"O campo de data não está no formato solicitado (dd/MM/yyyy), por favor reescreva.",
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

					// volta a janela principal para o estado anterior
					dispose();
				}
			}
		});
		btnSave.setBounds(149, 505, 129, 23);
		panel.add(btnSave);

		JLabel lblChamado = new JLabel("Chamado *");
		lblChamado.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblChamado.setBounds(15, 15, 93, 29);
		panel.add(lblChamado);

		JLabel lblNomeDoColaborador = new JLabel("Nome do colaborador *");
		lblNomeDoColaborador.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoColaborador.setBounds(124, 15, 288, 29);
		panel.add(lblNomeDoColaborador);

		JLabel lblUsurioDeRede = new JLabel("Usuário de rede *");
		lblUsurioDeRede.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblUsurioDeRede.setBounds(15, 68, 102, 29);
		panel.add(lblUsurioDeRede);

		JLabel lblDescrioDoItem = new JLabel("Descrição do item *");
		lblDescrioDoItem.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDescrioDoItem.setBounds(15, 121, 128, 29);
		panel.add(lblDescrioDoItem);

		JLabel lblCentroDeCusto = new JLabel("Centro de custo *");
		lblCentroDeCusto.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblCentroDeCusto.setBounds(124, 68, 288, 29);
		panel.add(lblCentroDeCusto);

		JLabel lblAtivo = new JLabel("Ativo *");
		lblAtivo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblAtivo.setBounds(15, 175, 99, 29);
		panel.add(lblAtivo);

		JLabel lblDispositivo = new JLabel("Dispositivo *");
		lblDispositivo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDispositivo.setBounds(118, 175, 97, 29);
		panel.add(lblDispositivo);

		JLabel lblHostname = new JLabel("Hostname *");
		lblHostname.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblHostname.setBounds(220, 175, 91, 29);
		panel.add(lblHostname);

		JLabel lblModelo = new JLabel("Modelo *");
		lblModelo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblModelo.setBounds(15, 236, 105, 29);
		panel.add(lblModelo);

		JLabel lblServiceTag = new JLabel("Service TAG *");
		lblServiceTag.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblServiceTag.setBounds(177, 237, 129, 29);
		panel.add(lblServiceTag);

		JLabel lblEspecificaesDoProcessador = new JLabel("Especificações do processador *");
		lblEspecificaesDoProcessador.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblEspecificaesDoProcessador.setBounds(15, 286, 189, 29);
		panel.add(lblEspecificaesDoProcessador);

		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico (elaborador do laudo) *");
		lblNomeDoTcnico.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoTcnico.setBounds(15, 391, 237, 29);
		panel.add(lblNomeDoTcnico);

		comboBoxDispositivo = new JComboBox();
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
		comboBoxDispositivo.setBounds(121, 198, 94, 29);
		panel.add(comboBoxDispositivo);
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
}
