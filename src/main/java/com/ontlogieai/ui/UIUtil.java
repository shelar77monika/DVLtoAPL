package com.ontlogieai.ui;

import com.ontlogieai.file.FileProcessor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;

public class UIUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(UIUtil.class);

    private final FileProcessor fileProcessor;

    public UIUtil(){
        fileProcessor = new FileProcessor();
    }

    public void setLookAndFeel() {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            LOGGER.info("Look and feel set successfully.");
        } catch (Exception e) {
            LOGGER.error("Error setting look and feel", e);
        }
    }

    public void createAndShowGUI() {
        JFrame frame = new JFrame("DVL TO AVL");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 300);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        JLabel headerLabel = createLabel("Company Name (TODO)", Font.BOLD, 20);
        JButton uploadButton = createButton("DVL To AVL");

        uploadButton.addActionListener(e -> fileProcessor.handleFileUpload(frame));

        mainPanel.add(headerLabel);
        mainPanel.add(Box.createRigidArea(new Dimension(0, 15)));
        mainPanel.add(uploadButton);

        frame.add(mainPanel);
        frame.setVisible(true);
    }

    private static JLabel createLabel(String text, int style, int size) {
        JLabel label = new JLabel(text);
        label.setFont(new Font("Arial", style, size));
        label.setAlignmentX(Component.CENTER_ALIGNMENT);
        return label;
    }

    private static JButton createButton(String text) {
        JButton button = new JButton(text);
        button.setAlignmentX(Component.CENTER_ALIGNMENT);
        button.setFont(new Font("Arial", Font.PLAIN, 16));
        button.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(Color.LIGHT_GRAY),
                new EmptyBorder(10, 15, 10, 15)
        ));
        return button;
    }
}
