package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;

import com.formdev.flatlaf.FlatIntelliJLaf;
import com.formdev.flatlaf.FlatLaf;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;

@SuppressWarnings("serial")
public class Opcoes extends JDialog {

	private final JPanel contentPanel = new JPanel();
	private JTextField textField_3;
	private JTextField textField_4;
	private JTextField textField_1;
	private Image img_folder = new ImageIcon(SplashAnimation.class.getResource("/resources/folder.png")).getImage()
			.getScaledInstance(18, 15, Image.SCALE_SMOOTH);
	private JTextField textField_2;
	private JTextField textField_5;

	@SuppressWarnings("rawtypes")
	public Opcoes(JFrame parentFrame) {
		super(parentFrame, "Ferrramentas - Opções", true); // Janela modal
		// Shortcuts
		// Captura a linha 5 == 4
		String pathData = ConfigManager.getDirectoryFromConfigLine(4);
		// Captura a linha 4 == 3
		String pathLocalModel = ConfigManager.getDirectoryFromConfigLine(3);
		// Captura a linha 3 == 2
		String pathPdfGenerated = ConfigManager.getDirectoryFromConfigLine(2);
		// Captura a linha 2 == 1
		String pathBackup = ConfigManager.getDirectoryFromConfigLine(1);
		// Captura a linha 1 == 0
		String pathConfig = ConfigManager.getDirectoryFromConfigLine(0);

		// Desatualizado > File diretorioInicial = new File("\\\\fscatorg01\\...");
		// //Substituir o local do pdf

		setTitle("Ferrramentas - Opções");
		setBounds(100, 100, 696, 300);
		setLocationRelativeTo(null);

		getContentPane().setLayout(new BorderLayout());
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(contentPanel, BorderLayout.CENTER);
		contentPanel.setLayout(null);
		{

			JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
			tabbedPane.setFont(new Font("Dialog", Font.PLAIN, 12));
			tabbedPane.setBounds(10, 11, 660, 217);
			contentPanel.add(tabbedPane);
			{
				JPanel panel = new JPanel();


				// Crie uma instância personalizada de TitledBorder
				TitledBorder titledBorder = new TitledBorder(
						new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
						"Personalização", TitledBorder.LEADING, TitledBorder.TOP, new Font("Dialog", Font.PLAIN, 12));

				panel.setBorder(titledBorder);
				tabbedPane.addTab("Geral", null, panel, null);
				panel.setLayout(null);

				JLabel lblNewLabel = new JLabel("TEMAS... (Em construção)");
				lblNewLabel.setBounds(10, 53, 169, 14);
				lblNewLabel.setBackground(Color.RED);
				panel.add(lblNewLabel);
				{
					JComboBox comboBox = new JComboBox();
					comboBox.setBounds(10, 22, 155, 22);
					panel.add(comboBox);
				}

			}
			{
				JPanel panel = new JPanel();
				panel.setBorder(new TitledBorder(null, "Pastas de trabalho", TitledBorder.LEADING, TitledBorder.TOP,
						null, null));
				tabbedPane.addTab("Paths", null, panel, null);
				panel.setLayout(null);
				{
					JLabel lblNewLabel_1 = new JLabel("Data folder:");
					lblNewLabel_1.setFont(new Font("Tahoma", Font.PLAIN, 12));
					lblNewLabel_1.setBounds(10, 24, 139, 14);
					panel.add(lblNewLabel_1);
				}
				{
					JLabel lblNewLabel_2 = new JLabel("Backup Word folder:");
					lblNewLabel_2.setFont(new Font("Tahoma", Font.PLAIN, 12));
					lblNewLabel_2.setBounds(10, 52, 139, 14);
					panel.add(lblNewLabel_2);
				}
				{
					JLabel lblNewLabel_3 = new JLabel("Laudos PDF folder:");
					lblNewLabel_3.setFont(new Font("Tahoma", Font.PLAIN, 12));
					lblNewLabel_3.setBounds(10, 80, 139, 14);
					panel.add(lblNewLabel_3);
				}
				{
					JLabel lblNewLabel_4 = new JLabel("Local Modelo laudo:");
					lblNewLabel_4.setFont(new Font("Tahoma", Font.PLAIN, 12));
					lblNewLabel_4.setBounds(10, 108, 139, 14);
					panel.add(lblNewLabel_4);
				}
				{
					JLabel lblNewLabel_5 = new JLabel("Arquivo de configuração:");
					lblNewLabel_5.setFont(new Font("Tahoma", Font.PLAIN, 12));
					lblNewLabel_5.setBounds(10, 136, 139, 14);
					panel.add(lblNewLabel_5);
				}
				textField_1 = new JTextField();
				textField_1.setEditable(false);
				textField_1.setFont(new Font("Tahoma", Font.PLAIN, 11));
				textField_1.setColumns(10);
				textField_1.setBounds(159, 21, 456, 20);
				textField_1.setText(pathData);
				panel.add(textField_1);
				{
					textField_2 = new JTextField();
					textField_2.setFont(new Font("Tahoma", Font.PLAIN, 11));
					textField_2.setEditable(false);
					textField_2.setColumns(10);
					textField_2.setBounds(159, 49, 456, 20);
					textField_2.setText(pathBackup);
					panel.add(textField_2);
				}
				{
					textField_3 = new JTextField();
					textField_3.setEditable(false);
					textField_3.setFont(new Font("Tahoma", Font.PLAIN, 11));
					textField_3.setBounds(159, 77, 456, 20);
					textField_3.setColumns(10);
					textField_3.setText(pathPdfGenerated);
					panel.add(textField_3);
				}
				textField_4 = new JTextField();
				textField_4.setEditable(false);
				textField_4.setFont(new Font("Tahoma", Font.PLAIN, 11));
				textField_4.setColumns(10);
				textField_4.setBounds(159, 105, 456, 20);
				textField_4.setText(pathLocalModel);
				panel.add(textField_4);
				{
					textField_5 = new JTextField();
					textField_5.setFont(new Font("Tahoma", Font.PLAIN, 11));
					textField_5.setEditable(false);
					textField_5.setColumns(10);
					textField_5.setBounds(159, 133, 456, 20);
					textField_5.setText(pathConfig);
					panel.add(textField_5);
				}
				{
					JLabel lblNewLabel_1 = new JLabel("");
					lblNewLabel_1.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent e) {
							try {
								Desktop.getDesktop().open(new File(pathData));
							} catch (IOException e1) {
								e1.printStackTrace();
							} finally {
								new File(pathData).mkdirs();
								try {
									Desktop.getDesktop().open(new File(pathData));
								} catch (IOException e1) {
									// TODO Auto-generated catch block
									e1.printStackTrace();
								}
							}
						}

						@Override
						public void mouseEntered(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
						}

						@Override
						public void mouseExited(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
						}
					});
					lblNewLabel_1.setIcon(new ImageIcon(img_folder));
					lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_1.setBounds(625, 24, 20, 14);
					panel.add(lblNewLabel_1);
				}
				{
					JLabel lblNewLabel_2 = new JLabel("");
					lblNewLabel_2.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent arg0) {
							try {
								Desktop.getDesktop().open(new File(pathBackup));
							} catch (IOException e) {
								e.printStackTrace();
							} finally {
								new File(pathBackup).mkdirs();
								try {
									Desktop.getDesktop().open(new File(pathBackup));
								} catch (IOException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							}
						}

						@Override
						public void mouseEntered(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
						}

						@Override
						public void mouseExited(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
						}
					});
					lblNewLabel_2.setIcon(new ImageIcon(img_folder));
					lblNewLabel_2.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_2.setBounds(625, 52, 20, 14);
					panel.add(lblNewLabel_2);

				}
				{
					JLabel lblNewLabel_3 = new JLabel("");
					lblNewLabel_3.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent arg0) {
							try {
								Desktop.getDesktop().open(new File(pathPdfGenerated));
							} catch (IOException e) {
								e.printStackTrace();
							} finally {
								new File(pathPdfGenerated).mkdirs();
								try {
									Desktop.getDesktop().open(new File(pathPdfGenerated));
								} catch (IOException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							}
						}

						@Override
						public void mouseEntered(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
						}

						@Override
						public void mouseExited(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
						}
					});
					lblNewLabel_3.setIcon(new ImageIcon(img_folder));
					lblNewLabel_3.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_3.setBounds(625, 80, 20, 14);
					panel.add(lblNewLabel_3);
				}
				{
					JLabel lblNewLabel_4 = new JLabel("");
					lblNewLabel_4.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent e) {
							if (pathLocalModel.equals("N/A")) {
								JOptionPane.showMessageDialog(null,
										"Não há registros salvos, o usuário optou por realizar o download da template (laudo técnico) diretamente do github ao invés de abrir o arquivo (.docx) localmente.");
							} else {
								try {
									Desktop.getDesktop().open(new File(pathLocalModel).getParentFile());
								} catch (IOException e1) {
									// TODO Auto-generated catch block
									e1.printStackTrace();
								} finally {
									new File(pathLocalModel).mkdirs();
									try {
										Desktop.getDesktop().open(new File(pathLocalModel).getParentFile());
									} catch (IOException e1) {
										// TODO Auto-generated catch block
										e1.printStackTrace();
									}
								}
							}
						}

						@Override
						public void mouseEntered(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
						}

						@Override
						public void mouseExited(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
						}
					});
					lblNewLabel_4.setIcon(new ImageIcon(img_folder));
					lblNewLabel_4.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_4.setBounds(625, 108, 20, 14);
					panel.add(lblNewLabel_4);
				}
				{
					JLabel lblNewLabel_5 = new JLabel("");
					lblNewLabel_5.addMouseListener(new MouseAdapter() {
						@Override
						public void mouseClicked(MouseEvent e) {
							try {
								Desktop.getDesktop().open(new File(pathConfig).getParentFile());
							} catch (IOException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							} finally {
								new File(pathBackup).mkdirs();
								try {
									Desktop.getDesktop().open(new File(pathConfig).getParentFile());
								} catch (IOException e1) {
									// TODO Auto-generated catch block
									e1.printStackTrace();
								}
							}
						}

						@Override
						public void mouseEntered(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
						}

						@Override
						public void mouseExited(MouseEvent e) {
							setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
						}
					});
					lblNewLabel_5.setIcon(new ImageIcon(img_folder));
					lblNewLabel_5.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_5.setBounds(625, 136, 20, 14);
					panel.add(lblNewLabel_5);
				}
			}
		}
		{
			JPanel buttonPane = new JPanel();
			buttonPane.setLayout(new FlowLayout(FlowLayout.RIGHT));
			getContentPane().add(buttonPane, BorderLayout.SOUTH);
			{
				JButton okButton = new JButton("OK");
				okButton.addActionListener(new ActionListener() {
					public void actionPerformed(ActionEvent arg0) {
						dispose();
					}
				});
				okButton.setActionCommand("OK");
				buttonPane.add(okButton);
				getRootPane().setDefaultButton(okButton);
			}
			{
				JButton cancelButton = new JButton("Cancel");
				cancelButton.addActionListener(new ActionListener() {

					public void actionPerformed(ActionEvent arg0) {
						dispose();
					}
				});
				cancelButton.setActionCommand("Cancel");
				buttonPane.add(cancelButton);
			}
		}
	}

	public static void main(String[] args) {
		try {
			FlatLaf.registerCustomDefaultsSource("com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator");
			FlatMacDarkLaf.setup();
			JFrame frame = new JFrame(); // Crie um JFrame para usar como janela principal
			Opcoes dialog = new Opcoes(frame); // Passe o JFrame como argumento
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
			dialog.setVisible(true);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
