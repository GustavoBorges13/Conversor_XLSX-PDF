package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.EventQueue;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Rectangle;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Font;
import javax.swing.border.TitledBorder;
import javax.swing.border.EtchedBorder;
import java.awt.Color;
import java.awt.Component;
import javax.swing.JTextField;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JSpinner;
import javax.swing.JButton;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.ChangeEvent;
import javax.swing.SpinnerNumberModel;

@SuppressWarnings("rawtypes")
public class EditarPlanilha extends JFrame {

	private static final long serialVersionUID = -7565309808695712383L;
	private JPanel contentPane;
	static JTextField txtLaudo;
	static JTextField txtNomeSolicitante;
	static JTextField txtUsuario;
	static JTextField txtCentroDeCusto;
	static JTextField txtItem;
	static JTextField txtAtivo;
	static JTextField txtDispositivo;
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
	static Component c; //contro


	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
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
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 450, 632);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
		GraphicsDevice defaultScreen = ge.getDefaultScreenDevice();
		@SuppressWarnings("unused")
		Rectangle rect = defaultScreen.getDefaultConfiguration().getBounds();
		// rect.getMaxX() - getWidth();

		int tolerancia = -150;
		int x = 0;
		int y = ((getWidth()) / 2) + tolerancia;
		setLocation(x, y);
		setVisible(true);

		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel lblNewLabel = new JLabel("Editar informações da planilha");
		lblNewLabel.setFont(new Font("Tahoma", Font.PLAIN, 17));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBounds(10, 11, 414, 21);
		contentPane.add(lblNewLabel);

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				": : Modo edi\u00E7\u00E3o : :", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(10, 36, 414, 546);
		contentPane.add(panel);
		panel.setLayout(null);

		txtLaudo = new JTextField();
		txtLaudo.setForeground(Color.RED);
		txtLaudo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtLaudo.getText().equals("N° Laudo")) {
					txtLaudo.setText("");
					txtLaudo.setForeground(Color.BLACK);
				} else {
					txtLaudo.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtLaudo.getText().equals("")) {
					txtLaudo.setForeground(Color.RED);
					txtLaudo.setText("N° Laudo");
				}

			}
		});
		txtLaudo.setToolTipText("");
		txtLaudo.setText("N° Laudo");
		txtLaudo.setBounds(15, 37, 99, 29);
		panel.add(txtLaudo);
		txtLaudo.setColumns(10);

		txtNomeSolicitante = new JTextField();
		txtNomeSolicitante.setForeground(Color.RED);
		txtNomeSolicitante.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtNomeSolicitante.getText().equals("Nome do solicitante")) {
					txtNomeSolicitante.setText("");
					txtNomeSolicitante.setForeground(Color.BLACK);
				} else {
					txtNomeSolicitante.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtNomeSolicitante.getText().equals("")) {
					txtNomeSolicitante.setForeground(Color.RED);
					txtNomeSolicitante.setText("Nome do solicitante");
				}

			}
		});
		txtNomeSolicitante.setText("Nome do solicitante");
		txtNomeSolicitante.setColumns(10);
		txtNomeSolicitante.setBounds(124, 37, 276, 29);
		panel.add(txtNomeSolicitante);

		txtUsuario = new JTextField();
		txtUsuario.setForeground(Color.RED);
		txtUsuario.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtUsuario.getText().equals("Usuário")) {
					txtUsuario.setText("");
					txtUsuario.setForeground(Color.BLACK);
				} else {
					txtUsuario.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtUsuario.getText().equals("")) {
					txtUsuario.setForeground(Color.RED);
					txtUsuario.setText("Usuário");
				}

			}
		});
		txtUsuario.setText("Usuário");
		txtUsuario.setColumns(10);
		txtUsuario.setBounds(15, 90, 98, 29);
		panel.add(txtUsuario);

		txtCentroDeCusto = new JTextField();
		txtCentroDeCusto.setForeground(Color.RED);
		txtCentroDeCusto.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtCentroDeCusto.getText().equals("Centro de Custo")) {
					txtCentroDeCusto.setText("");
					txtCentroDeCusto.setForeground(Color.BLACK);
				} else {
					txtCentroDeCusto.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtCentroDeCusto.getText().equals("")) {
					txtCentroDeCusto.setForeground(Color.RED);
					txtCentroDeCusto.setText("Centro de Custo");
				}

			}
		});
		txtCentroDeCusto.setText("Centro de Custo");
		txtCentroDeCusto.setColumns(10);
		txtCentroDeCusto.setBounds(123, 90, 277, 29);
		panel.add(txtCentroDeCusto);

		txtItem = new JTextField();
		txtItem.setForeground(Color.RED);
		txtItem.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtItem.getText().equals("Item")) {
					txtItem.setText("");
					txtItem.setForeground(Color.BLACK);
				} else {
					txtItem.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtItem.getText().equals("")) {
					txtItem.setForeground(Color.RED);
					txtItem.setText("Item");
				}

			}
		});
		txtItem.setText("Item");
		txtItem.setColumns(10);
		txtItem.setBounds(15, 144, 291, 29);
		panel.add(txtItem);

		comboBoxQuantidade = new JComboBox();
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
		comboBoxQuantidade
				.setModel(new DefaultComboBoxModel(new String[] {"Selecionar", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"}));
		comboBoxQuantidade.setBounds(311, 144, 89, 29);
		panel.add(comboBoxQuantidade);

		JLabel lblQTD = new JLabel("Quantidade:");
		lblQTD.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblQTD.setBounds(311, 119, 89, 29);
		panel.add(lblQTD);

		txtAtivo = new JTextField();
		txtAtivo.setForeground(Color.RED);
		txtAtivo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtAtivo.getText().equals("Ativo")) {
					txtAtivo.setText("");
					txtAtivo.setForeground(Color.BLACK);
				} else {
					txtAtivo.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtAtivo.getText().equals("")) {
					txtAtivo.setForeground(Color.RED);
					txtAtivo.setText("Ativo");
				}

			}
		});
		txtAtivo.setText("Ativo");
		txtAtivo.setColumns(10);
		txtAtivo.setBounds(15, 198, 106, 29);
		panel.add(txtAtivo);

		txtDispositivo = new JTextField();
		txtDispositivo.setForeground(Color.RED);
		txtDispositivo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtDispositivo.getText().equals("Dispositivo")) {
					txtDispositivo.setText("");
					txtDispositivo.setForeground(Color.BLACK);
				} else {
					txtDispositivo.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtDispositivo.getText().equals("")) {
					txtDispositivo.setForeground(Color.RED);
					txtDispositivo.setText("Dispositivo");
				}

			}
		});
		txtDispositivo.setText("Dispositivo");
		txtDispositivo.setColumns(10);
		txtDispositivo.setBounds(128, 198, 69, 29);
		panel.add(txtDispositivo);

		txtHostname = new JTextField();
		txtHostname.setForeground(Color.RED);
		txtHostname.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtHostname.getText().equals("Hostname")) {
					txtHostname.setText("");
					txtHostname.setForeground(Color.BLACK);
				} else {
					txtHostname.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtHostname.getText().equals("")) {
					txtHostname.setForeground(Color.RED);
					txtHostname.setText("Hostname");
				}

			}
		});
		txtHostname.setText("Hostname");
		txtHostname.setColumns(10);
		txtHostname.setBounds(204, 198, 102, 29);
		panel.add(txtHostname);

		comboBoxFabricante = new JComboBox();
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
		comboBoxFabricante.setBounds(311, 198, 89, 29);
		panel.add(comboBoxFabricante);

		lblFabricante = new JLabel("Fabricante:");
		lblFabricante.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblFabricante.setBounds(311, 175, 89, 29);
		panel.add(lblFabricante);

		txtModelo = new JTextField();
		txtModelo.setForeground(Color.RED);
		txtModelo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtModelo.getText().equals("Modelo")) {
					txtModelo.setText("");
					txtModelo.setForeground(Color.BLACK);
				} else {
					txtModelo.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtModelo.getText().equals("")) {
					txtModelo.setForeground(Color.RED);
					txtModelo.setText("Modelo");
				}

			}
		});
		txtModelo.setText("Modelo");
		txtModelo.setColumns(10);
		txtModelo.setBounds(15, 258, 152, 29);
		panel.add(txtModelo);

		txtServiceTag = new JTextField();
		txtServiceTag.setForeground(Color.RED);
		txtServiceTag.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtServiceTag.getText().equals("Service TAG")) {
					txtServiceTag.setText("");
					txtServiceTag.setForeground(Color.BLACK);
				} else {
					txtServiceTag.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtServiceTag.getText().equals("")) {
					txtServiceTag.setForeground(Color.RED);
					txtServiceTag.setText("Service TAG");
				}

			}
		});
		txtServiceTag.setText("Service TAG");
		txtServiceTag.setColumns(10);
		txtServiceTag.setBounds(177, 258, 129, 29);
		panel.add(txtServiceTag);

		txtDdmmyyyy = new JTextField();
		txtDdmmyyyy.setForeground(Color.RED);
		txtDdmmyyyy.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtDdmmyyyy.getText().equals("dd/mm/yyyy")) {
					txtDdmmyyyy.setText("");
					txtDdmmyyyy.setForeground(Color.BLACK);
				} else {
					txtDdmmyyyy.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtDdmmyyyy.getText().equals("")) {
					txtDdmmyyyy.setForeground(Color.RED);
					txtDdmmyyyy.setText("dd/mm/yyyy");
				}

			}
		});
		txtDdmmyyyy.setText("dd/mm/yyyy");
		txtDdmmyyyy.setColumns(10);
		txtDdmmyyyy.setBounds(309, 258, 91, 29);
		panel.add(txtDdmmyyyy);

		lblDataAquisicao = new JLabel("Data aquisição:");
		lblDataAquisicao.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDataAquisicao.setBounds(309, 236, 99, 29);
		panel.add(lblDataAquisicao);

		txtCpu = new JTextField();
		txtCpu.setForeground(Color.RED);
		txtCpu.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtCpu.getText().equals("CPU")) {
					txtCpu.setText("");
					txtCpu.setForeground(Color.BLACK);
				} else {
					txtCpu.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtCpu.getText().equals("")) {
					txtCpu.setForeground(Color.RED);
					txtCpu.setText("CPU");
				}

			}
		});
		txtCpu.setText("CPU");
		txtCpu.setColumns(10);
		txtCpu.setBounds(15, 309, 385, 29);
		panel.add(txtCpu);

		lblsStorage = new JLabel("Storage (GB):");
		lblsStorage.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblsStorage.setBounds(15, 337, 83, 29);
		panel.add(lblsStorage);

		JLabel lblMemoria = new JLabel("Memória (GB):");
		lblMemoria.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblMemoria.setBounds(124, 337, 79, 29);
		panel.add(lblMemoria);

		spinner_memoria = new JSpinner();
		spinner_memoria.setModel(new SpinnerNumberModel(Integer.valueOf(0), Integer.valueOf(0), null, Integer.valueOf(1)));
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

		spinner_memoria.setBounds(124, 362, 99, 29);
		panel.add(spinner_memoria);

		txtNomeDoTecnico = new JTextField();
		txtNomeDoTecnico.setForeground(Color.RED);
		txtNomeDoTecnico.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtNomeDoTecnico.getText().equals("Nome do técnico")) {
					txtNomeDoTecnico.setText("");
					txtNomeDoTecnico.setForeground(Color.BLACK);
				} else {
					txtNomeDoTecnico.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtNomeDoTecnico.getText().equals("")) {
					txtNomeDoTecnico.setForeground(Color.RED);
					txtNomeDoTecnico.setText("Nome do técnico");
				}

			}
		});
		txtNomeDoTecnico.setText("Nome do técnico");
		txtNomeDoTecnico.setColumns(10);
		txtNomeDoTecnico.setBounds(15, 415, 385, 29);
		panel.add(txtNomeDoTecnico);

		txtObservao = new JTextField();
		txtObservao.setForeground(Color.RED);
		txtObservao.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtObservao.getText().equals("N/D ou Nosso")) {
					txtObservao.setText("");
				} else {
					txtObservao.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtObservao.getText().equals("")) {
					txtObservao.setForeground(Color.BLACK);
					txtObservao.setText("N/D ou Nosso");
				}

			}
		});
		txtObservao.setColumns(10);
		txtObservao.setBounds(15, 466, 190, 29);
		panel.add(txtObservao);

		txtStatus = new JTextField();
		txtStatus.setForeground(Color.RED);
		txtStatus.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if (txtStatus.getText().equals("N/D")) {
					txtStatus.setText("");
				} else {
					txtStatus.selectAll();
				}
			}

			@Override
			public void focusLost(FocusEvent e) {
				if (txtStatus.getText().equals("")) {
					txtStatus.setForeground(Color.BLACK);
					txtStatus.setText("N/D");
				}

			}
		});
		txtStatus.setColumns(10);
		txtStatus.setBounds(220, 466, 180, 29);
		panel.add(txtStatus);

		lblObservacao = new JLabel("Observação:");
		lblObservacao.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblObservacao.setBounds(15, 442, 85, 29);
		panel.add(lblObservacao);

		lblStatus = new JLabel("Status:");
		lblStatus.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblStatus.setBounds(220, 442, 83, 29);
		panel.add(lblStatus);

		JButton btnSave = new JButton("Salvar");
		btnSave.setBounds(146, 512, 129, 23);
		panel.add(btnSave);

		comboBoxStorage = new JComboBox();
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
		comboBoxStorage.setModel(new DefaultComboBoxModel(new String[] { "Selecionar", "250 HD", "300 HD", "500 HD",
				"750 HD", "1000 HD", "120 SSD", "240 SSD" }));
		comboBoxStorage.setBounds(15, 362, 99, 29);
		panel.add(comboBoxStorage);
		
		JLabel lblChamado = new JLabel("Chamado:");
		lblChamado.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblChamado.setBounds(15, 15, 93, 29);
		panel.add(lblChamado);
		
		JLabel lblNomeDoColaborador = new JLabel("Nome do colaborador:");
		lblNomeDoColaborador.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoColaborador.setBounds(124, 15, 147, 29);
		panel.add(lblNomeDoColaborador);
		
		JLabel lblUsurioDeRede = new JLabel("Usuário de rede:");
		lblUsurioDeRede.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblUsurioDeRede.setBounds(15, 68, 102, 29);
		panel.add(lblUsurioDeRede);
		
		JLabel lblDescrioDoItem = new JLabel("Descrição do item:");
		lblDescrioDoItem.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDescrioDoItem.setBounds(15, 121, 128, 29);
		panel.add(lblDescrioDoItem);
		
		JLabel lblCentroDeCusto = new JLabel("Centro de custo:");
		lblCentroDeCusto.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblCentroDeCusto.setBounds(124, 68, 195, 29);
		panel.add(lblCentroDeCusto);
		
		JLabel lblAtivo = new JLabel("Ativo:");
		lblAtivo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblAtivo.setBounds(15, 175, 106, 29);
		panel.add(lblAtivo);
		
		JLabel lblDispositivo = new JLabel("Dispositivo:");
		lblDispositivo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblDispositivo.setBounds(128, 175, 69, 29);
		panel.add(lblDispositivo);
		
		JLabel lblHostname = new JLabel("Hostname:");
		lblHostname.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblHostname.setBounds(204, 175, 102, 29);
		panel.add(lblHostname);
		
		JLabel lblModelo = new JLabel("Modelo:");
		lblModelo.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblModelo.setBounds(15, 236, 105, 29);
		panel.add(lblModelo);
		
		JLabel lblServiceTag = new JLabel("Service TAG:");
		lblServiceTag.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblServiceTag.setBounds(177, 237, 129, 29);
		panel.add(lblServiceTag);
		
		JLabel lblEspecificaesDoProcessador = new JLabel("Especificações do processador:");
		lblEspecificaesDoProcessador.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblEspecificaesDoProcessador.setBounds(15, 286, 189, 29);
		panel.add(lblEspecificaesDoProcessador);
		
		JLabel lblNomeDoTcnico = new JLabel("Nome do técnico (elaborador do laudo):");
		lblNomeDoTcnico.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNomeDoTcnico.setBounds(15, 391, 237, 29);
		panel.add(lblNomeDoTcnico);
		requestFocus();
	}
}
