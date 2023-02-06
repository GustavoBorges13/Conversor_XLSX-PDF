package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.TitledBorder;
import com.formdev.flatlaf.FlatIntelliJLaf;
import javax.swing.JComboBox;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.ImageIcon;
import java.awt.Font;

@SuppressWarnings("serial")
public class Opcoes extends JDialog {

	private final JPanel contentPanel = new JPanel();
	private JTextField textField_3;
	private JTextField textField_2;
	private JTextField textField_1;
	private Image img_folder = new ImageIcon(SplashAnimation.class.getResource("/resources/folder.png")).getImage()
			.getScaledInstance(18, 15, Image.SCALE_SMOOTH);
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		try {
			UIManager.setLookAndFeel(new FlatIntelliJLaf());
			Opcoes dialog = new Opcoes();
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
			dialog.setVisible(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public Opcoes() {
		setTitle("Ferrramentas - Opções");
		setBounds(100, 100, 500, 300);
		setLocationRelativeTo(null);
		setVisible(true);
		getContentPane().setLayout(new BorderLayout());
		contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(contentPanel, BorderLayout.CENTER);
		contentPanel.setLayout(null);
		{
			JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
			tabbedPane.setBackground(new Color(240, 240, 240));
			tabbedPane.setBounds(10, 11, 464, 217);
			contentPanel.add(tabbedPane);
			{
				JPanel panel = new JPanel();
				panel.setBackground(Color.WHITE);
				panel.setBorder(new TitledBorder(new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)), "Personaliza\u00E7\u00E3o", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
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
				panel.setBorder(new TitledBorder(null, "Pastas de trabalho", TitledBorder.LEADING, TitledBorder.TOP, null, null));
				tabbedPane.addTab("Paths", null, panel, null);
				panel.setLayout(null);
				{
					JLabel lblNewLabel_1 = new JLabel("Pasta de DataSource:");
					lblNewLabel_1.setBounds(10, 24, 139, 14);
					panel.add(lblNewLabel_1);
				}
				{
					JLabel lblNewLabel_2 = new JLabel("Pasta de Backup (Word):");
					lblNewLabel_2.setBounds(10, 48, 139, 14);
					panel.add(lblNewLabel_2);
				}
				{
					JLabel lblNewLabel_2 = new JLabel("Pasta de laudos em PDF:");
					lblNewLabel_2.setBounds(10, 73, 139, 14);
					panel.add(lblNewLabel_2);
				}
				{
					textField_3 = new JTextField();
					textField_3.setEditable(false);
					textField_3.setFont(new Font("Tahoma", Font.PLAIN, 10));
					textField_3.setBounds(146, 70, 273, 20);
					panel.add(textField_3);
					textField_3.setColumns(10);
				}
				{
					textField_2 = new JTextField();
					textField_2.setEditable(false);
					textField_2.setFont(new Font("Tahoma", Font.PLAIN, 10));
					textField_2.setColumns(10);
					textField_2.setBounds(146, 45, 273, 20);
					panel.add(textField_2);
				}
				{
					textField_1 = new JTextField();
					textField_1.setEditable(false);
					textField_1.setFont(new Font("Tahoma", Font.PLAIN, 10));
					textField_1.setColumns(10);
					textField_1.setBounds(146, 21, 273, 20);
					panel.add(textField_1);
				}
				{
					JLabel lblNewLabel_1 = new JLabel("");
					lblNewLabel_1.setIcon(new ImageIcon(img_folder));
					lblNewLabel_1.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_1.setBounds(429, 25, 20, 14);
					panel.add(lblNewLabel_1);
				}
				{
					JLabel lblNewLabel_2 = new JLabel("");
					lblNewLabel_2.setIcon(new ImageIcon(img_folder));
					lblNewLabel_2.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_2.setBounds(429, 49, 20, 14);
					panel.add(lblNewLabel_2);
				}
				{
					JLabel lblNewLabel_3 = new JLabel("");
					lblNewLabel_3.setIcon(new ImageIcon(img_folder));
					lblNewLabel_3.setHorizontalAlignment(SwingConstants.CENTER);
					lblNewLabel_3.setBounds(429, 74, 20, 14);
					panel.add(lblNewLabel_3);
				}
			}
		}
		{
			JPanel buttonPane = new JPanel();
			buttonPane.setLayout(new FlowLayout(FlowLayout.RIGHT));
			getContentPane().add(buttonPane, BorderLayout.SOUTH);
			{
				JButton okButton = new JButton("OK");
				okButton.setActionCommand("OK");
				buttonPane.add(okButton);
				getRootPane().setDefaultButton(okButton);
			}
			{
				JButton cancelButton = new JButton("Cancel");
				cancelButton.addActionListener(new ActionListener() {

					public void actionPerformed(ActionEvent arg0) {
						System.exit(0);
					}
				});
				cancelButton.setActionCommand("Cancel");
				buttonPane.add(cancelButton);
			}
		}
	}
}
