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
import javax.swing.JTextField;
import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JSpinner;
import javax.swing.JButton;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;

public class EditarPlanilha extends JFrame {

	private static final long serialVersionUID = -7565309808695712383L;
	private JPanel contentPane;
	private JTextField txtLaudo;
	private JTextField txtNomeSolicitante;
	private JTextField txtUsuario;
	private JTextField txtCentroDeCusto;
	private JTextField txtItem;
	private JTextField txtAtivo;
	private JTextField txtDispositivo;
	private JTextField txtHostname;
	private JComboBox comboBoxFabricante;
	private JLabel lblFabricante;
	private JTextField txtModelo;
	private JTextField txtServiceTag;
	private JTextField txtDdmmyyyy;
	private JLabel lblDataAquisicao;
	private JTextField txtCpu;
	private JLabel lblsStorage;
	private JTextField txtNomeDoTecnico;
	private JTextField txtObservao;
	private JTextField txtStatus;
	private JLabel lblObservacao;
	private JLabel lblStatus;
	private JComboBox comboBoxStorage;
	private JLabel lblNewLabel_1;

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

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public EditarPlanilha() {
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 450, 632);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
		//GraphicsDevice defaultScreen = ge.getDefaultScreenDevice();
		//Rectangle rect = defaultScreen.getDefaultConfiguration().getBounds();
		//rect.getMaxX() - getWidth()
		
		int tolerancia = 25;
		int x = getHeight()/4;
		int y = (getWidth()/2)-tolerancia;
		setLocation(x, y);
		setVisible(true);
		
		
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JLabel lblNewLabel = new JLabel("Editar  informações da planilha");
		lblNewLabel.setFont(new Font("Tahoma", Font.PLAIN, 17));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel.setBounds(10, 11, 414, 21);
		contentPane.add(lblNewLabel);
		
		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)), ": : Modo edi\u00E7\u00E3o : :", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel.setBounds(10, 43, 414, 539);
		contentPane.add(panel);
		panel.setLayout(null);
		
		txtLaudo = new JTextField();
		txtLaudo.setForeground(Color.RED);
		txtLaudo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtLaudo.getText().equals("N° Laudo")) {
					txtLaudo.setText("");
					txtLaudo.setForeground(Color.BLACK);
				}else {
					txtLaudo.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtLaudo.getText().equals("")) {
					txtLaudo.setForeground(Color.RED);
					txtLaudo.setText("N° Laudo");
				}
				
			}
		});
		txtLaudo.setToolTipText("");
		txtLaudo.setText("N° Laudo");
		txtLaudo.setBounds(10, 23, 99, 29);
		panel.add(txtLaudo);
		txtLaudo.setColumns(10);
		
		txtNomeSolicitante = new JTextField();
		txtNomeSolicitante.setForeground(Color.RED);
		txtNomeSolicitante.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtNomeSolicitante.getText().equals("Nome do solicitante")) {
					txtNomeSolicitante.setText("");
					txtNomeSolicitante.setForeground(Color.BLACK);
				}else {
					txtNomeSolicitante.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtNomeSolicitante.getText().equals("")) {
					txtNomeSolicitante.setForeground(Color.RED);
					txtNomeSolicitante.setText("Nome do solicitante");
				}
				
			}
		});
		txtNomeSolicitante.setText("Nome do solicitante");
		txtNomeSolicitante.setColumns(10);
		txtNomeSolicitante.setBounds(119, 23, 285, 29);
		panel.add(txtNomeSolicitante);
		
		txtUsuario = new JTextField();
		txtUsuario.setForeground(Color.RED);
		txtUsuario.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtUsuario.getText().equals("Usuário")) {
					txtUsuario.setText("");
					txtUsuario.setForeground(Color.BLACK);
				}else {
					txtUsuario.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtUsuario.getText().equals("")) {
					txtUsuario.setForeground(Color.RED);
					txtUsuario.setText("Usuário");
				}
				
			}
		});
		txtUsuario.setText("Usuário");
		txtUsuario.setColumns(10);
		txtUsuario.setBounds(10, 63, 99, 29);
		panel.add(txtUsuario);
		
		txtCentroDeCusto = new JTextField();
		txtCentroDeCusto.setForeground(Color.RED);
		txtCentroDeCusto.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtCentroDeCusto.getText().equals("Centro de Custo")) {
					txtCentroDeCusto.setText("");
					txtCentroDeCusto.setForeground(Color.BLACK);
				}else {
					txtCentroDeCusto.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtCentroDeCusto.getText().equals("")) {
					txtCentroDeCusto.setForeground(Color.RED);
					txtCentroDeCusto.setText("Centro de Custo");
				}
				
			}
		});
		txtCentroDeCusto.setText("Centro de Custo");
		txtCentroDeCusto.setColumns(10);
		txtCentroDeCusto.setBounds(119, 63, 285, 29);
		panel.add(txtCentroDeCusto);
		
		txtItem = new JTextField();
		txtItem.setForeground(Color.RED);
		txtItem.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtItem.getText().equals("Item")) {
					txtItem.setText("");
					txtItem.setForeground(Color.BLACK);
				}else {
					txtItem.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtItem.getText().equals("")) {
					txtItem.setForeground(Color.RED);
					txtItem.setText("Item");
				}
				
			}
		});
		txtItem.setText("Item");
		txtItem.setColumns(10);
		txtItem.setBounds(10, 124, 285, 29);
		panel.add(txtItem);
		
		JComboBox comboBoxQuantidade = new JComboBox();
		comboBoxQuantidade.setMaximumRowCount(10);
		comboBoxQuantidade.setModel(new DefaultComboBoxModel(new String[] {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10"}));
		comboBoxQuantidade.setBounds(305, 124, 99, 29);
		panel.add(comboBoxQuantidade);
		
		JLabel lblQTD = new JLabel("Quantidade:");
		lblQTD.setBounds(305, 99, 99, 29);
		panel.add(lblQTD);
		
		txtAtivo = new JTextField();
		txtAtivo.setForeground(Color.RED);
		txtAtivo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtAtivo.getText().equals("Ativo")) {
					txtAtivo.setText("");
					txtAtivo.setForeground(Color.BLACK);
				}else {
					txtAtivo.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtAtivo.getText().equals("")) {
					txtAtivo.setForeground(Color.RED);
					txtAtivo.setText("Ativo");
				}
				
			}
		});
		txtAtivo.setText("Ativo");
		txtAtivo.setColumns(10);
		txtAtivo.setBounds(10, 180, 99, 29);
		panel.add(txtAtivo);
		
		txtDispositivo = new JTextField();
		txtDispositivo.setForeground(Color.RED);
		txtDispositivo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtDispositivo.getText().equals("Dispositivo")) {
					txtDispositivo.setText("");
					txtDispositivo.setForeground(Color.BLACK);
				}else {
					txtDispositivo.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtDispositivo.getText().equals("")) {
					txtDispositivo.setForeground(Color.RED);
					txtDispositivo.setText("Dispositivo");
				}
				
			}
		});
		txtDispositivo.setText("Dispositivo");
		txtDispositivo.setColumns(10);
		txtDispositivo.setBounds(119, 180, 80, 29);
		panel.add(txtDispositivo);
		
		txtHostname = new JTextField();
		txtHostname.setForeground(Color.RED);
		txtHostname.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtHostname.getText().equals("Hostname")) {
					txtHostname.setText("");
					txtHostname.setForeground(Color.BLACK);
				}else {
					txtHostname.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtHostname.getText().equals("")) {
					txtHostname.setForeground(Color.RED);
					txtHostname.setText("Hostname");
				}
				
			}
		});
		txtHostname.setText("Hostname");
		txtHostname.setColumns(10);
		txtHostname.setBounds(209, 180, 86, 29);
		panel.add(txtHostname);
		
		comboBoxFabricante = new JComboBox();
		comboBoxFabricante.setModel(new DefaultComboBoxModel(new String[] {"Selecionar", "Dell Inc", "Lenovo"}));
		comboBoxFabricante.setMaximumRowCount(2);
		comboBoxFabricante.setBounds(305, 180, 99, 29);
		panel.add(comboBoxFabricante);
		
		lblFabricante = new JLabel("Fabricante:");
		lblFabricante.setBounds(305, 157, 99, 29);
		panel.add(lblFabricante);
		
		txtModelo = new JTextField();
		txtModelo.setForeground(Color.RED);
		txtModelo.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtModelo.getText().equals("Modelo")) {
					txtModelo.setText("");
					txtModelo.setForeground(Color.BLACK);
				}else {
					txtModelo.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtModelo.getText().equals("")) {
					txtModelo.setForeground(Color.RED);
					txtModelo.setText("Modelo");
				}
				
			}
		});
		txtModelo.setText("Modelo");
		txtModelo.setColumns(10);
		txtModelo.setBounds(11, 234, 153, 29);
		panel.add(txtModelo);
		
		txtServiceTag = new JTextField();
		txtServiceTag.setForeground(Color.RED);
		txtServiceTag.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtServiceTag.getText().equals("Service TAG")) {
					txtServiceTag.setText("");
					txtServiceTag.setForeground(Color.BLACK);
				}else {
					txtServiceTag.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtServiceTag.getText().equals("")) {
					txtServiceTag.setForeground(Color.RED);
					txtServiceTag.setText("Service TAG");
				}
				
			}
		});
		txtServiceTag.setText("Service TAG");
		txtServiceTag.setColumns(10);
		txtServiceTag.setBounds(174, 234, 122, 29);
		panel.add(txtServiceTag);
		
		txtDdmmyyyy = new JTextField();
		txtDdmmyyyy.setForeground(Color.RED);
		txtDdmmyyyy.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtDdmmyyyy.getText().equals("dd/mm/yyyy")) {
					txtDdmmyyyy.setText("");
					txtDdmmyyyy.setForeground(Color.BLACK);
				}else {
					txtDdmmyyyy.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtDdmmyyyy.getText().equals("")) {
					txtDdmmyyyy.setForeground(Color.RED);
					txtDdmmyyyy.setText("dd/mm/yyyy");
				}
				
			}
		});
		txtDdmmyyyy.setText("dd/mm/yyyy");
		txtDdmmyyyy.setColumns(10);
		txtDdmmyyyy.setBounds(306, 234, 99, 29);
		panel.add(txtDdmmyyyy);
		
		lblDataAquisicao = new JLabel("Data aquisição:");
		lblDataAquisicao.setBounds(306, 212, 99, 29);
		panel.add(lblDataAquisicao);
		
		txtCpu = new JTextField();
		txtCpu.setForeground(Color.RED);
		txtCpu.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtCpu.getText().equals("CPU")) {
					txtCpu.setText("");
					txtCpu.setForeground(Color.BLACK);
				}else {
					txtCpu.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtCpu.getText().equals("")) {
					txtCpu.setForeground(Color.RED);
					txtCpu.setText("CPU");
				}
				
			}
		});
		txtCpu.setText("CPU");
		txtCpu.setColumns(10);
		txtCpu.setBounds(11, 274, 394, 29);
		panel.add(txtCpu);
		
		lblsStorage = new JLabel("Storage (GB):");
		lblsStorage.setBounds(11, 308, 83, 29);
		panel.add(lblsStorage);
		
		JLabel lblMemoria = new JLabel("Memória (GB):");
		lblMemoria.setBounds(116, 308, 83, 29);
		panel.add(lblMemoria);
		
		JSpinner spinner_memoria = new JSpinner();
		spinner_memoria.setBounds(117, 333, 83, 29);
		panel.add(spinner_memoria);
		
		txtNomeDoTecnico = new JTextField();
		txtNomeDoTecnico.setForeground(Color.RED);
		txtNomeDoTecnico.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtNomeDoTecnico.getText().equals("Nome do técnico")) {
					txtNomeDoTecnico.setText("");
					txtNomeDoTecnico.setForeground(Color.BLACK);
				}else {
					txtNomeDoTecnico.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtNomeDoTecnico.getText().equals("")) {
					txtNomeDoTecnico.setForeground(Color.RED);
					txtNomeDoTecnico.setText("Nome do técnico");
				}
				
			}
		});
		txtNomeDoTecnico.setText("Nome do técnico");
		txtNomeDoTecnico.setColumns(10);
		txtNomeDoTecnico.setBounds(10, 387, 394, 29);
		panel.add(txtNomeDoTecnico);
		
		txtObservao = new JTextField();
		txtObservao.setForeground(Color.RED);
		txtObservao.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtObservao.getText().equals("N/D ou Nosso")) {
					txtObservao.setText("");
					txtObservao.setForeground(Color.BLACK);
				}else {
					txtObservao.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtObservao.getText().equals("")) {
					txtObservao.setForeground(Color.RED);
					txtObservao.setText("N/D ou Nosso");
				}
				
			}
		});
		txtObservao.setText("N/D ou Nosso");
		txtObservao.setColumns(10);
		txtObservao.setBounds(11, 452, 125, 29);
		panel.add(txtObservao);
		
		txtStatus = new JTextField();
		txtStatus.setForeground(Color.RED);
		txtStatus.addFocusListener(new FocusAdapter() {
			@Override
			public void focusGained(FocusEvent arg0) {
				if(txtStatus.getText().equals("N/D")) {
					txtStatus.setText("");
					txtStatus.setForeground(Color.BLACK);
				}else {
					txtStatus.selectAll();
				}
			}
			@Override
			public void focusLost(FocusEvent e) {
				if(txtStatus.getText().equals("")) {
					txtStatus.setForeground(Color.RED);
					txtStatus.setText("N/D");
				}
				
			}
		});
		txtStatus.setText("N/D");
		txtStatus.setColumns(10);
		txtStatus.setBounds(147, 452, 99, 29);
		panel.add(txtStatus);
		
		lblObservacao = new JLabel("Observação");
		lblObservacao.setBounds(10, 427, 83, 29);
		panel.add(lblObservacao);
		
		lblStatus = new JLabel("Status");
		lblStatus.setBounds(146, 424, 83, 29);
		panel.add(lblStatus);
		
		JButton btnSave = new JButton("Salvar");
		btnSave.setBounds(14, 504, 129, 23);
		panel.add(btnSave);
		
		JButton btnAdicionar = new JButton("Adicionar");
		btnAdicionar.setBounds(155, 504, 129, 23);
		panel.add(btnAdicionar);
		
		comboBoxStorage = new JComboBox();
		comboBoxStorage.setModel(new DefaultComboBoxModel(new String[] {"Selecionar", "250 HD", "300 HD", "500 HD", "750 HD", "1000 HD", "120 SSD", "240 SSD"}));
		comboBoxStorage.setMaximumRowCount(10);
		comboBoxStorage.setBounds(11, 333, 99, 29);
		panel.add(comboBoxStorage);
		
		lblNewLabel_1 = new JLabel("---------- Finalização ----------");
		lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
		lblNewLabel_1.setBounds(8, 367, 394, 23);
		panel.add(lblNewLabel_1);
	}
}
