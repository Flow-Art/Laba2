/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package laba2;

import java.util.logging.Logger;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;

/**
 *
 * @author Admin
 */
public class MainFrame extends javax.swing.JFrame {

    /**
     * Creates new form MainFrame
     */
    public MainFrame() {
        initComponents();
        
    }
    ExcelManipulator MyEM = new ExcelManipulator();
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        FileWasNotImportedExceptionDialog = new javax.swing.JDialog();
        CloseExcFrameButton = new javax.swing.JButton();
        ExcLabel = new javax.swing.JLabel();
        ExpCompDialog = new javax.swing.JDialog();
        ExpCompButton = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        ImpCompDialog = new javax.swing.JDialog();
        jLabel2 = new javax.swing.JLabel();
        ImpCompButton = new javax.swing.JButton();
        ExportButton = new javax.swing.JButton();
        ResultButton = new javax.swing.JButton();
        ExitButton = new javax.swing.JButton();

        CloseExcFrameButton.setText("OK");
        CloseExcFrameButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CloseExcFrameButtonActionPerformed(evt);
            }
        });

        ExcLabel.setText("Файл не был импортирован! Невозможно экспортировать результаты.");

        javax.swing.GroupLayout FileWasNotImportedExceptionDialogLayout = new javax.swing.GroupLayout(FileWasNotImportedExceptionDialog.getContentPane());
        FileWasNotImportedExceptionDialog.getContentPane().setLayout(FileWasNotImportedExceptionDialogLayout);
        FileWasNotImportedExceptionDialogLayout.setHorizontalGroup(
            FileWasNotImportedExceptionDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(FileWasNotImportedExceptionDialogLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(ExcLabel)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(FileWasNotImportedExceptionDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(FileWasNotImportedExceptionDialogLayout.createSequentialGroup()
                    .addGap(175, 175, 175)
                    .addComponent(CloseExcFrameButton)
                    .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );
        FileWasNotImportedExceptionDialogLayout.setVerticalGroup(
            FileWasNotImportedExceptionDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(FileWasNotImportedExceptionDialogLayout.createSequentialGroup()
                .addGap(74, 74, 74)
                .addComponent(ExcLabel)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(FileWasNotImportedExceptionDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(FileWasNotImportedExceptionDialogLayout.createSequentialGroup()
                    .addGap(137, 137, 137)
                    .addComponent(CloseExcFrameButton)
                    .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );

        ExpCompButton.setText("OK");
        ExpCompButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ExpCompButtonActionPerformed(evt);
            }
        });

        jLabel1.setText("Экспорт выполнен успешно!");

        javax.swing.GroupLayout ExpCompDialogLayout = new javax.swing.GroupLayout(ExpCompDialog.getContentPane());
        ExpCompDialog.getContentPane().setLayout(ExpCompDialogLayout);
        ExpCompDialogLayout.setHorizontalGroup(
            ExpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(ExpCompDialogLayout.createSequentialGroup()
                .addGroup(ExpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(ExpCompDialogLayout.createSequentialGroup()
                        .addGap(149, 149, 149)
                        .addComponent(ExpCompButton))
                    .addGroup(ExpCompDialogLayout.createSequentialGroup()
                        .addGap(86, 86, 86)
                        .addComponent(jLabel1)))
                .addContainerGap(147, Short.MAX_VALUE))
        );
        ExpCompDialogLayout.setVerticalGroup(
            ExpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, ExpCompDialogLayout.createSequentialGroup()
                .addContainerGap(126, Short.MAX_VALUE)
                .addComponent(jLabel1)
                .addGap(18, 18, 18)
                .addComponent(ExpCompButton)
                .addGap(115, 115, 115))
        );

        jLabel2.setText("Импорт выполнен успешно!");

        ImpCompButton.setText("OK");
        ImpCompButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ImpCompButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout ImpCompDialogLayout = new javax.swing.GroupLayout(ImpCompDialog.getContentPane());
        ImpCompDialog.getContentPane().setLayout(ImpCompDialogLayout);
        ImpCompDialogLayout.setHorizontalGroup(
            ImpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(ImpCompDialogLayout.createSequentialGroup()
                .addGroup(ImpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(ImpCompDialogLayout.createSequentialGroup()
                        .addGap(103, 103, 103)
                        .addComponent(jLabel2))
                    .addGroup(ImpCompDialogLayout.createSequentialGroup()
                        .addGap(159, 159, 159)
                        .addComponent(ImpCompButton)))
                .addContainerGap(134, Short.MAX_VALUE))
        );
        ImpCompDialogLayout.setVerticalGroup(
            ImpCompDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(ImpCompDialogLayout.createSequentialGroup()
                .addGap(84, 84, 84)
                .addComponent(jLabel2)
                .addGap(18, 18, 18)
                .addComponent(ImpCompButton)
                .addContainerGap(157, Short.MAX_VALUE))
        );

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        ExportButton.setText("Импортировать файл");
        ExportButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ExportButtonActionPerformed(evt);
            }
        });

        ResultButton.setText("Экспортировать результаты расчётов");
        ResultButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ResultButtonActionPerformed(evt);
            }
        });

        ExitButton.setText("Выйти из программы");
        ExitButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ExitButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(96, 96, 96)
                        .addComponent(ExportButton))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(53, 53, 53)
                        .addComponent(ResultButton))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(100, 100, 100)
                        .addComponent(ExitButton)))
                .addContainerGap(90, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(42, 42, 42)
                .addComponent(ExportButton)
                .addGap(18, 18, 18)
                .addComponent(ResultButton)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(ExitButton)
                .addContainerGap(152, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void ExportButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ExportButtonActionPerformed
        try {
            MyEM.export();
            ImpCompDialog.setSize(500, 250);
            ImpCompDialog.setVisible(true);
        } catch (IOException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
        }
       
    }//GEN-LAST:event_ExportButtonActionPerformed

    private void ResultButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ResultButtonActionPerformed
        try {
            MyEM.result();
            ExpCompDialog.setSize(500, 250);
            ExpCompDialog.setVisible(true);
        } catch (IOException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
        } catch (FileWasNotImportedException ex) {   
            FileWasNotImportedExceptionDialog.setSize(500, 250);
            FileWasNotImportedExceptionDialog.setVisible(true);
        }
    }//GEN-LAST:event_ResultButtonActionPerformed

    private void ExitButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ExitButtonActionPerformed
        System.exit(0);
    }//GEN-LAST:event_ExitButtonActionPerformed

    private void CloseExcFrameButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CloseExcFrameButtonActionPerformed
        FileWasNotImportedExceptionDialog.setVisible(false);
    }//GEN-LAST:event_CloseExcFrameButtonActionPerformed

    private void ExpCompButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ExpCompButtonActionPerformed
        ExpCompDialog.setVisible(false);
    }//GEN-LAST:event_ExpCompButtonActionPerformed

    private void ImpCompButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ImpCompButtonActionPerformed
        ImpCompDialog.setVisible(false);
    }//GEN-LAST:event_ImpCompButtonActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton CloseExcFrameButton;
    private javax.swing.JLabel ExcLabel;
    private javax.swing.JButton ExitButton;
    private javax.swing.JButton ExpCompButton;
    private javax.swing.JDialog ExpCompDialog;
    private javax.swing.JButton ExportButton;
    private javax.swing.JDialog FileWasNotImportedExceptionDialog;
    private javax.swing.JButton ImpCompButton;
    private javax.swing.JDialog ImpCompDialog;
    private javax.swing.JButton ResultButton;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    // End of variables declaration//GEN-END:variables
}
