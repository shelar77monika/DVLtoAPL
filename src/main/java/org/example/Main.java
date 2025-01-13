package org.example;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

public class Main {
    private static final String UPLOAD_DIR = "uploads/";

    public static void main(String[] args) {
        // Ensure the upload directory exists
        File uploadDir = new File(UPLOAD_DIR);
        if (!uploadDir.exists()) {
            uploadDir.mkdirs();
        }

        // Set Look and Feel
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Create the GUI
        JFrame frame = new JFrame("DVL TO AVL");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 300);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
        mainPanel.setBorder(new EmptyBorder(20, 20, 20, 20));

        // Header Label
        JLabel headerLabel = new JLabel("DVL (Device/Instrument List) to Alarm Parameter List");
        headerLabel.setFont(new Font("Arial", Font.BOLD, 20));
        headerLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        headerLabel.setBorder(new EmptyBorder(10, 0, 20, 0));

        // Buttons
        JButton uploadButton = new JButton("Upload File");
        JButton downloadButton = new JButton("Download File");
        uploadButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        downloadButton.setAlignmentX(Component.CENTER_ALIGNMENT);
        uploadButton.setFont(new Font("Arial", Font.PLAIN, 16));
        downloadButton.setFont(new Font("Arial", Font.PLAIN, 16));

        // Status Label
        JLabel statusLabel = new JLabel("Status: Ready");
        statusLabel.setFont(new Font("Arial", Font.ITALIC, 14));
        statusLabel.setAlignmentX(Component.CENTER_ALIGNMENT);
        statusLabel.setForeground(Color.DARK_GRAY);
        statusLabel.setBorder(new EmptyBorder(20, 0, 10, 0));

        // Add spacing and borders
        uploadButton.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(Color.LIGHT_GRAY),
                new EmptyBorder(10, 15, 10, 15)
        ));
        downloadButton.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(Color.LIGHT_GRAY),
                new EmptyBorder(10, 15, 10, 15)
        ));

        // File Upload Logic
        uploadButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int result = fileChooser.showOpenDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                String status = uploadFile(selectedFile);
                statusLabel.setText("Status: " + status);
            }
        });

        // File Download Logic
        downloadButton.addActionListener(e -> {
            File[] files = uploadDir.listFiles();
            if (files == null || files.length == 0) {
                JOptionPane.showMessageDialog(frame, "No files available to download.", "Info", JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            // Let the user select a file from the uploads directory
            String[] fileNames = new String[files.length];
            for (int i = 0; i < files.length; i++) {
                fileNames[i] = files[i].getName();
            }
            String selectedFile = (String) JOptionPane.showInputDialog(
                    frame,
                    "Select a file to download:",
                    "Download File",
                    JOptionPane.PLAIN_MESSAGE,
                    null,
                    fileNames,
                    fileNames[0]);

            if (selectedFile != null) {
                JFileChooser saveChooser = new JFileChooser();
                saveChooser.setSelectedFile(new File(selectedFile));
                int result = saveChooser.showSaveDialog(frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File saveFile = saveChooser.getSelectedFile();
                    String status = downloadFile(selectedFile, saveFile);
                    statusLabel.setText("Status: " + status);
                }
            }
        });

        // Add components to main panel
        mainPanel.add(headerLabel);
        mainPanel.add(uploadButton);
        mainPanel.add(Box.createRigidArea(new Dimension(0, 15))); // Spacer
        mainPanel.add(downloadButton);
        mainPanel.add(Box.createRigidArea(new Dimension(0, 20))); // Spacer
        mainPanel.add(statusLabel);

        // Add panel to frame and display
        frame.add(mainPanel);
        frame.setVisible(true);
    }

    private static String uploadFile(File file) {
        try {
            Path destination = Paths.get(UPLOAD_DIR + file.getName());
            Files.copy(file.toPath(), destination, StandardCopyOption.REPLACE_EXISTING);
            return "File uploaded successfully: " + file.getName();
        } catch (IOException e) {
            e.printStackTrace();
            return "Failed to upload file: " + e.getMessage();
        }
    }

    private static String downloadFile(String fileName, File destination) {
        try {
            Path source = Paths.get(UPLOAD_DIR + fileName);
            Files.copy(source, destination.toPath(), StandardCopyOption.REPLACE_EXISTING);
            return "File downloaded successfully: " + destination.getName();
        } catch (IOException e) {
            e.printStackTrace();
            return "Failed to download file: " + e.getMessage();
        }
    }
}
