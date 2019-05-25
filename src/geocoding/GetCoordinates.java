/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package geocoding;

import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.event.AdjustmentEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author gianlucalovecchio
 */
public final class GetCoordinates extends javax.swing.JFrame {

    private String path = "";

    private ArrayList<String> addressList = new ArrayList();
    private static boolean toRun = true;
    private static int currentApi = 1;

    public GetCoordinates() {
        initComponents();
        jRadioButton1.setSelected(true);
        jRadioButton1.setEnabled(true);
        jRadioButton2.setEnabled(true);
        jRadioButton3.setEnabled(true);
        jLabel1.setText(getApiKey(1));
        this.setResizable(false);
        file.setPreferredSize(new Dimension(100, 100));
        file.setOpaque(true);
        setTableVisible(table, false);
        jButton5.setEnabled(false);
        if (getApiKey(currentApi).equals("")) {
            jButton1.setEnabled(false);
            jButton2.setEnabled(false);
            jButton3.setEnabled(false);
        }

    }

    public void setTableVisible(JTable table, boolean isVisible) {
        table.setVisible(isVisible);
        table.getTableHeader().setVisible(isVisible);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jMenuItem1 = new javax.swing.JMenuItem();
        jMenuItem2 = new javax.swing.JMenuItem();
        menuBar1 = new java.awt.MenuBar();
        menu1 = new java.awt.Menu();
        menu2 = new java.awt.Menu();
        buttonGroup1 = new javax.swing.ButtonGroup();
        jButton1 = new javax.swing.JButton();
        file = new javax.swing.JLabel();
        jButton2 = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        table = new javax.swing.JTable();
        jButton3 = new javax.swing.JButton();
        jButton4 = new javax.swing.JButton();
        pbar = new javax.swing.JProgressBar();
        status = new javax.swing.JLabel();
        jRadioButton1 = new javax.swing.JRadioButton();
        jRadioButton2 = new javax.swing.JRadioButton();
        jRadioButton3 = new javax.swing.JRadioButton();
        jLabel1 = new javax.swing.JLabel();
        jButton5 = new javax.swing.JButton();

        jMenuItem1.setText("jMenuItem1");

        jMenuItem2.setText("jMenuItem2");

        menu1.setLabel("File");
        menuBar1.add(menu1);

        menu2.setLabel("Edit");
        menuBar1.add(menu2);

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jButton1.setText("Seleccionar archivo Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        file.setBackground(new java.awt.Color(204, 204, 204));
        file.setForeground(new java.awt.Color(0, 102, 255));
        file.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jButton2.setText("Obtener Coordenadas");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        table.setAutoCreateRowSorter(true);
        table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jScrollPane1.setViewportView(table);

        jButton3.setText("Generar Excel");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jButton4.setIcon(new javax.swing.ImageIcon(getClass().getResource("/geocoding/img/key.png"))); // NOI18N
        jButton4.setText("Añadir API Key");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        status.setText(" ");

        buttonGroup1.add(jRadioButton1);
        jRadioButton1.setText("API Google");
        jRadioButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jRadioButton1ActionPerformed(evt);
            }
        });

        buttonGroup1.add(jRadioButton2);
        jRadioButton2.setText("API Bing");
        jRadioButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jRadioButton2ActionPerformed(evt);
            }
        });

        buttonGroup1.add(jRadioButton3);
        jRadioButton3.setText("API Yandex");
        jRadioButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jRadioButton3ActionPerformed(evt);
            }
        });

        jLabel1.setBackground(new java.awt.Color(204, 204, 204));
        jLabel1.setBorder(javax.swing.BorderFactory.createEtchedBorder());

        jButton5.setText("STOP");
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jRadioButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(72, 72, 72)
                        .addComponent(jRadioButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(66, 66, 66)
                        .addComponent(jRadioButton3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jScrollPane1)
                            .addComponent(pbar, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(file, javax.swing.GroupLayout.PREFERRED_SIZE, 512, javax.swing.GroupLayout.PREFERRED_SIZE)
                                            .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 512, javax.swing.GroupLayout.PREFERRED_SIZE))
                                        .addGap(4, 4, 4))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 190, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(72, 72, 72)
                                        .addComponent(jButton3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGap(64, 64, 64)))
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jButton5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(jButton4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 190, Short.MAX_VALUE))))
                        .addContainerGap())
                    .addComponent(status, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(15, 15, 15)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 29, Short.MAX_VALUE)
                    .addComponent(jButton4, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(file, javax.swing.GroupLayout.DEFAULT_SIZE, 29, Short.MAX_VALUE)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton2)
                    .addComponent(jButton3)
                    .addComponent(jButton5))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 9, Short.MAX_VALUE)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jRadioButton2)
                    .addComponent(jRadioButton3)
                    .addComponent(jRadioButton1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(status, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(2, 2, 2)
                .addComponent(pbar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 399, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("EXCEL FILES", new String[]{"xlsx"});
        fileChooser.setFileFilter(filter);
        fileChooser.setVisible(true);
        fileChooser.addChoosableFileFilter(filter);
        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            String str = selectedFile.getAbsolutePath();
            String temp = "";
            if (str != null && str.length() > 100) {
                temp = str.substring(0, 100) + "...";
            } else {
                temp = str;
            }
            String labelText = String.format("<html>" + temp + "</html>");
            file.setText(labelText);
            path = str;
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        if (!path.equals("")) {
            (t = new aTask()).execute();
        } else {
            JOptionPane.showMessageDialog(this, "No se ha seleccionado ningun archivo", "Error", JOptionPane.ERROR_MESSAGE);
        }

    }//GEN-LAST:event_jButton2ActionPerformed

    private class aTask extends SwingWorker<Void, String> {

        @Override
        protected Void doInBackground() throws Exception {
            getLocations();
            Thread.sleep(100);
            return null;
        }
    }

    public String createURL(String address) {
        String url = "";
        switch (currentApi) {
            case 1:
                url = "https://maps.googleapis.com/maps/api/geocode/xml?address=" + address + "&key=" + getApiKey(currentApi);
                break;
            case 2:
                url = "http://dev.virtualearth.net/REST/v1/Locations?q=" + address + "&key=" + getApiKey(currentApi);
                break;
            case 3:
                if ("".equals(getApiKey(currentApi))) {
                    url = "https://geocode-maps.yandex.ru/1.x/?geocode=" + address + "&lang=en_US";
                    jLabel1.setText("API Key isn't set, using free version of Yandex API");

                } else {
                    url = "https://geocode-maps.yandex.ru/1.x/?lang=en_US&apikey=" + getApiKey(currentApi) + "&geocode=" + address;
                }
                break;
            default:
                break;
        }
        return url;
    }

    public void enableButtons(boolean state) {
        jButton1.setEnabled(state);
        jButton2.setEnabled(state);
        jButton3.setEnabled(state);
        jButton4.setEnabled(state);
        jRadioButton1.setEnabled(state);
        jRadioButton2.setEnabled(state);
        jRadioButton3.setEnabled(state);

    }

    public void getLocations() throws InterruptedException, MalformedURLException, IOException, Exception {
        jButton5.setEnabled(true);

        addressList.clear();
        status.setText("<html>Cargando Excel... <small>(Puede tardar unos minutos)</small></html>");
        pbar.setMinimum(0);
        pbar.setMaximum(100);
        pbar.setStringPainted(true);
        pbar.setValue(0);
        enableButtons(false);
        DefaultTableModel modelo = new DefaultTableModel();
        modelo.setColumnIdentifiers(new Object[]{"Dirección", "Latitud", "Longitud"});
        setTableVisible(table, true);
        setCursor(new Cursor(WAIT_CURSOR));
        table.setModel(modelo);
//        jScrollPane1.getVerticalScrollBar().addAdjustmentListener((AdjustmentEvent e) -> {
//            e.getAdjustable().setValue(e.getAdjustable().getMaximum());
//        });

        FileInputStream fileInputStream = new FileInputStream(path);
        try {
            toRun = true;
            XSSFWorkbook wb = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = wb.getSheetAt(0);
            int rowTotal = worksheet.getLastRowNum();

            int cellTotal = 0;
            boolean error = false;
            String address;
            String coord = "";
            String lat = "";
            String lng = "";
            int percent = 0;
            status.setText("Obteniendo coordenadas");
            jButton5.setEnabled(true);
            OUTER:
            for (int i = rowTotal; i >= 0; i--) {
                if (toRun == true) {

                    DataFormatter formatter = new DataFormatter();
                    percent++;

                    if (worksheet.getRow(i) != null) {
                        XSSFRow row = worksheet.getRow(i);
                        cellTotal = row.getLastCellNum();

                        address = "";

                        for (int x = 0; x < cellTotal; x++) {

                            XSSFCell cell = row.getCell((short) x);
                            if (address.equals("")) {
                                address = formatter.formatCellValue(cell);
                            } else {
                                address = address + ", " + formatter.formatCellValue(cell);
                            }

                        }
                        int intIndex = 0;
                        String oldAdress = address;
                        address = URLEncoder.encode(address, "UTF-8");
                        String request = createURL(address);
                        URL url = new URL(request);
                        InputStream input = url.openStream();
                        String xml = IOUtils.toString(input, "UTF-8");
                        input.close();
                        switch (currentApi) {
                            case 1:
                                intIndex = xml.indexOf("OVER_QUERY_LIMIT");
                                if (intIndex != - 1) {
                                    JOptionPane.showMessageDialog(this, "You have exceeded your daily request quota for this API.", "Error", JOptionPane.ERROR_MESSAGE);
                                    error = true;
                                    status.setText("<html><span style=\"color:red\">Error: You have exceeded your daily request quota for this API.</span> </html>");
                                    break OUTER;
                                }
                                intIndex = xml.indexOf("ZERO_RESULTS");
                                if (intIndex == - 1) {

                                    coord = matchString("<location>", "</location>", xml);
                                    lat = matchString("<lat>", "</lat>", coord);
                                    lng = matchString("<lng>", "</lng>", coord);
                                    address = oldAdress + " ; " + lat + " ; " + lng;
                                    addressList.add(address);
                                } else {
                                    address = oldAdress + " ; Error ; Error";
                                    addressList.add(oldAdress + " ; Error ; Error");
                                }
                                break;
                            case 2:
                                intIndex = xml.indexOf("RateLimitExceeded");
                                if (intIndex != - 1) {
                                    JOptionPane.showMessageDialog(this, "You have exceeded your daily request quota for this API.", "Error", JOptionPane.ERROR_MESSAGE);
                                    error = true;
                                    status.setText("<html><span style=\"color:red\">Error: You have exceeded your daily request quota for this API.</span> </html>");
                                    break OUTER;
                                }
                                intIndex = xml.indexOf("InvalidRequest");
                                if (intIndex == - 1) {
                                    coord = matchString("\"coordinates\":", "}", xml);
                                    lat = matchString("[", ",", coord);
                                    lng = matchString(",", "]", coord);
                                    address = oldAdress + " ; " + lat + " ; " + lng;
                                    addressList.add(address);
                                } else {
                                    address = oldAdress + " ; Error ; Error";
                                    addressList.add(oldAdress + " ; Error ; Error");
                                }
                                break;

                            case 3:
                                intIndex = xml.indexOf("Hit rate limit");
                                if (intIndex != - 1) {
                                    JOptionPane.showMessageDialog(this, "You have exceeded your daily request quota for this API.", "Error", JOptionPane.ERROR_MESSAGE);
                                    error = true;
                                    status.setText("<html><span style=\"color:red\">Error: You have exceeded your daily request quota for this API.</span> </html>");
                                    break OUTER;
                                }
                                intIndex = xml.indexOf("<found>0</found>");
                                if (intIndex == - 1) {
                                    coord = matchString("<pos>", "</pos>", xml);
                                    String[] latlon = coord.split(" ");
                                    lat = latlon[1];
                                    lng = latlon[0];
                                    address = oldAdress + " ; " + lat + " ; " + lng;
                                    addressList.add(address);
                                } else {
                                    address = oldAdress + " ; Error ; Error";
                                    addressList.add(oldAdress + " ; Error ; Error");
                                }
                                break;
                            default:
                                break;
                        }
                        String[] split = address.split(" ; ");
                        modelo.insertRow(0, split);
                    }
                    pbar.setValue((int) ((percent * 100.0f) / rowTotal));
                } else {
                    enableButtons(true);
                    setCursor(new Cursor(DEFAULT_CURSOR));

                    jButton5.setEnabled(false);
                    i = -2;
                }

            }

            enableButtons(true);

            setCursor(new Cursor(DEFAULT_CURSOR));
            if (error == false && toRun == true) {
                JOptionPane.showMessageDialog(GetCoordinates.this, "Las direcciones se han procesado correctamente!", "", JOptionPane.INFORMATION_MESSAGE);
            }

        } catch (Exception e) {

            JOptionPane.showMessageDialog(GetCoordinates.this, "Error: " + e, "Error", JOptionPane.ERROR_MESSAGE);

            enableButtons(true);
            setCursor(new Cursor(DEFAULT_CURSOR));
        }
    }

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        try {
            if (path.equals("")) {
                JOptionPane.showMessageDialog(this, "No se ha seleccionado ningun archivo", "Error", JOptionPane.ERROR_MESSAGE);
            } else if (addressList == null) {
                (t = new aTask()).execute();
                writeFile(addressList);
            } else {
                writeFile(addressList);
            }
        } catch (IOException ex) {
            Logger.getLogger(GetCoordinates.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        try {
            AddApiKey api = new AddApiKey();
            api.setVisible(true);
            jButton1.setEnabled(true);
            jButton2.setEnabled(true);
            jButton3.setEnabled(true);
        } catch (IOException ex) {
            Logger.getLogger(GetCoordinates.class.getName()).log(Level.SEVERE, null, ex);
        }

    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        // TODO add your handling code here:
        toRun = false;
        jButton5.setEnabled(false);
        jRadioButton1.setEnabled(true);
        jRadioButton2.setEnabled(true);
        jRadioButton3.setEnabled(true);
    }//GEN-LAST:event_jButton5ActionPerformed

    private void jRadioButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jRadioButton2ActionPerformed
        // TODO add your handling code here:
        currentApi = 2;
        jLabel1.setText(getApiKey(currentApi));
    }//GEN-LAST:event_jRadioButton2ActionPerformed

    private void jRadioButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jRadioButton1ActionPerformed
        // TODO add your handling code here:

        currentApi = 1;
        jLabel1.setText(getApiKey(currentApi));
    }//GEN-LAST:event_jRadioButton1ActionPerformed

    private void jRadioButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jRadioButton3ActionPerformed
        // TODO add your handling code here:
        currentApi = 3;

        if ("".equals(getApiKey(currentApi))) {
            jLabel1.setText("API Key isn't set, using free version of Yandex API");
        } else {
            jLabel1.setText(getApiKey(currentApi));
        }
    }//GEN-LAST:event_jRadioButton3ActionPerformed

    public static String matchString(String startsWith, String endsWith, String file) {
        String requiredString = file.substring(file.indexOf(startsWith) + startsWith.length(), file.indexOf(endsWith));
        return requiredString;
    }

    public static void writeFile(ArrayList<String> finalAddress) throws IOException {
        Collections.reverse(finalAddress);
        finalAddress.add(0, "null");
        JFrame parentFrame = new JFrame();
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Specify a file to save");
        int userSelection = fileChooser.showSaveDialog(parentFrame);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File archivoXLS = fileChooser.getSelectedFile();
            if (FilenameUtils.getExtension(archivoXLS.getName()).equalsIgnoreCase("xlsx")) {
            } else {
                archivoXLS = new File(archivoXLS.toString() + ".xlsx");
                archivoXLS = new File(archivoXLS.getParentFile(), FilenameUtils.getBaseName(archivoXLS.getName()) + ".xlsx");
            }
            if (archivoXLS.exists()) {
                archivoXLS.delete();
            }
            archivoXLS.createNewFile();
            XSSFWorkbook libro = new XSSFWorkbook();
            FileOutputStream archivo = new FileOutputStream(archivoXLS);
            Sheet hoja = libro.createSheet();
            for (int f = 0; f < finalAddress.size(); f++) {
                Row fila = hoja.createRow(f);
                String[] split = finalAddress.get(f).split(" ; ");
                String[] header = new String[]{"Dirección", "Latitud", "Longitud"};
                for (int c = 0; c < 3; c++) {
                    Cell celda = fila.createCell(c);
                    if (f == 0) {
                        celda.setCellValue(header[c]);
                    } else {
                        celda.setCellValue(split[c]);
                    }
                }
            }
            libro.write(archivo);
            archivo.close();
            Desktop.getDesktop().open(archivoXLS);
        }
    }

    public static String getApiKey(int api) {
        Properties prop = new Properties();
        InputStream input = null;
        String key = "";
        try {
            input = new FileInputStream("config.properties");
            prop.load(input);
            switch (api) {
                case 1:
                    key = prop.getProperty("APIKeyGoogle");

                    break;
                case 2:
                    key = prop.getProperty("APIKeyBing");
                    break;
                case 3:
                    key = prop.getProperty("APIKeyYandex");
                    break;
                default:
                    break;
            }

        } catch (IOException ex) {
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                }
            }
        }
        return key;
    }

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
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(GetCoordinates.class
                    .getName()).log(java.util.logging.Level.SEVERE, null, ex);

        }
        //</editor-fold>

        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> {
            new GetCoordinates().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.ButtonGroup buttonGroup1;
    private javax.swing.JLabel file;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem2;
    private javax.swing.JRadioButton jRadioButton1;
    private javax.swing.JRadioButton jRadioButton2;
    private javax.swing.JRadioButton jRadioButton3;
    private javax.swing.JScrollPane jScrollPane1;
    private java.awt.Menu menu1;
    private java.awt.Menu menu2;
    private java.awt.MenuBar menuBar1;
    private javax.swing.JProgressBar pbar;
    private javax.swing.JLabel status;
    private javax.swing.JTable table;
    // End of variables declaration//GEN-END:variables
    private aTask t;
}
