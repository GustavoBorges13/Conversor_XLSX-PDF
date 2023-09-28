package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;

import swing.LiquidProgress;

import java.awt.FlowLayout;
import java.awt.BorderLayout;
import javax.swing.BoxLayout;
import java.awt.Scrollbar;
import javax.swing.JSlider;
import javax.swing.event.ChangeListener;
import javax.swing.event.ChangeEvent;

public class Main2 extends JFrame {

	private static final long serialVersionUID = 1L;
	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Main2 frame = new Main2();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Main2() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBounds(100, 100, 946, 602);
        contentPane = new JPanel();

        setContentPane(contentPane);
        contentPane.setLayout(null);
        
     // Cria uma instância de LiquidProgress
        LiquidProgress liquidProgress = new LiquidProgress();
        liquidProgress.setValue(50);

     // Defina o tamanho e a posição do LiquidProgress
        liquidProgress.setBounds(286, 107, 349, 290); // Ajuste os valores conforme necessário

       
        // Adicione o LiquidProgress ao JPanel
        contentPane.add(liquidProgress);
        
        JSlider slider = new JSlider();
        slider.addChangeListener(new ChangeListener() {
        	public void stateChanged(ChangeEvent arg0) {
        		liquidProgress.setValue(slider.getValue());
        	}
        });
        slider.setBounds(617, 526, 270, 26);
        contentPane.add(slider);
	}
}
