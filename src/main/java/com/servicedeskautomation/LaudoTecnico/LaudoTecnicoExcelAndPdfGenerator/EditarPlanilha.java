package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.EventQueue;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.Rectangle;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

public class EditarPlanilha extends JFrame {

	private static final long serialVersionUID = -7565309808695712383L;
	private JPanel contentPane;

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
	}

}
