package org.example;
import java.time.LocalDate;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.List;

public class simpleWindow extends JFrame {
    private  JPanel panel = new JPanel();
    simpleWindow(){
        super("Test window");
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        final boolean[] btnBol = {false,false};
        final JButton [] downBtnTXT = new JButton[] { new JButton("Записать в файл"),
                new JButton("Дописать в файл"),
                new JButton("Чтение из файла")};

        final JButton [] downBtnDOCX = new JButton[] { new JButton("Записать в файл"),
                new JButton("Дописать в файл"),
                new JButton("Чтение из файла")};

        final JButton [] downBtnXLS = new JButton[] { new JButton("Записать в файл"),
                new JButton("Дописать в файл"),
                new JButton("Чтение из файла")};

        final JButton [] downBtnCSV = new JButton[] { new JButton("Записать в файл"),
                new JButton("Дописать в файл"),
                new JButton("Чтение из файла")};


        final JButton [] buttonsCreateFile = new JButton[]{ new JButton("Формат TXT"),
                new JButton("Формат Word"),
                new JButton("Формат Excel"),
                new JButton("Формат CSV")};

        final JButton [] buttonsWriteFile = new JButton[]{ new JButton("Формат TXT"),
                new JButton("Формат Word"),
                new JButton("Формат Excel"),
                new JButton("Формат CSV")};

        final JButton [] buttonsMenu = new JButton[]{ new JButton(" Создание файлов "),
                new JButton(" Работа с файлами "),
                new JButton(" Просмотр отчетов "),
                new JButton(" Администрирование "),
                new JButton(" Выход ")};

        initHorizontal(buttonsMenu,20,20,220,true,200,50);
        initVetical(buttonsCreateFile,20,90,70,false,200,50);
        initVetical(buttonsWriteFile,240,90,70,false,200,50);
        initVetical(downBtnTXT,450,90,30,false,200,25);
        initVetical(downBtnDOCX,450,180,30,false,200,25);
        initVetical(downBtnXLS,450,270,30,false,200,25);
        initVetical(downBtnCSV,450,360,30,false,200,25);
        panel.setLayout(null);


        buttonsMenu[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                if(buttonsCreateFile[0].isVisible()){
                    editVisible(buttonsCreateFile,false);

                }
                else {
                    editlocation(buttonsCreateFile, 20, 90, 70);
                    editVisible(buttonsCreateFile, true);
                    btnBol[0]=true;
                }
            }
        });
        buttonsMenu[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                if(buttonsWriteFile[0].isVisible()){
                    editVisible(buttonsWriteFile,false);
                    editVisible(downBtnTXT,false);
                    editVisible(downBtnDOCX,false);
                }
                else {
                    editlocation(buttonsWriteFile, 240, 90, 70);
                    editVisible(buttonsWriteFile, true);
                    btnBol[1]=true;
                }

            }
        });
        buttonsMenu[4].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
            }
        });

        buttonsWriteFile[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                if(downBtnTXT[0].isVisible()){
                    editVisible(downBtnTXT,false);
                }
                else {

                    editVisible(downBtnTXT,true);
                }
            }
        });
        buttonsWriteFile[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                if(downBtnDOCX[0].isVisible()){
                    editVisible(downBtnDOCX,false);
                }
                else {

                    editVisible(downBtnDOCX,true);
                }
            }
        });
        buttonsWriteFile[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {


                if(downBtnXLS[0].isVisible()){
                    editVisible(downBtnXLS,false);
                }
                else {

                    editVisible(downBtnXLS,true);
                }
            }
        });
        buttonsWriteFile[3].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                if(downBtnCSV[0].isVisible()){
                    editVisible(downBtnCSV,false);
                }
                else {

                    editVisible(downBtnCSV,true);
                }
            }
        });

        // Создание файлов
        buttonsCreateFile[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    FileWriter fileWriter = new FileWriter("files/exampleTXT.txt");
                    JOptionPane.showMessageDialog(null, "Файл TXT успешно создан!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());;
                }
            }
        });
        buttonsCreateFile[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    FileWriter fileWriter = new FileWriter("files/exampleDOCX.docx");
                    JOptionPane.showMessageDialog(null, "Файл DOCX успешно создан!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());;
                }
            }
        });
        buttonsCreateFile[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    FileWriter fileWriter = new FileWriter("files/exampleXLS.xls");
                    JOptionPane.showMessageDialog(null, "Файл XLS успешно создан!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());;
                }
            }
        });
        buttonsCreateFile[3].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    FileWriter fileWriter = new FileWriter("files/exampleCSV.csv");
                    JOptionPane.showMessageDialog(null, "Файл CSV успешно создан!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());;
                }
            }
        });

        //Запись новых файлов

        downBtnTXT[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");

                String filePath = "files/exampleTXT.txt"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    // Создаем объект FileOutputStream для записи в файл
                    FileOutputStream fileOutputStream = new FileOutputStream(filePath);

                    // Записываем данные в файл
                    byte[] byteData = enteredText.getBytes();
                    fileOutputStream.write(byteData);

                    // Закрываем поток вывода файла
                    fileOutputStream.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        downBtnDOCX[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");

                String filePath = "files/exampleDOCX.docx"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    XWPFDocument document = new XWPFDocument();
                    XWPFParagraph paragraph = document.createParagraph();
                    XWPFRun run = paragraph.createRun();

                    // Создаем объект FileOutputStream для записи в файл
                    run.setText(enteredText);

                    FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                    document.write(fileOutputStream);
                    fileOutputStream.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        downBtnXLS[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");
                String enteredText1 = JOptionPane.showInputDialog(null, "колонка:");
                String enteredText2 = JOptionPane.showInputDialog(null, "строка:");
                String filePath = "files/exampleXLS.xls"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {

                    HSSFWorkbook workbook = new HSSFWorkbook();
                    Sheet sheet = workbook.createSheet("Sheet1");
                    Row row = sheet.createRow( Integer.parseInt(enteredText2)-1);
                    Cell cell = row.createCell(Integer.parseInt(enteredText1)-1);


                    cell.setCellValue(enteredText);
                    // Создаем объект FileOutputStream для записи в файл

                    FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                    workbook.write(fileOutputStream);
                    fileOutputStream.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                }
                catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        downBtnCSV[0].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст в формате CSV(text,text,text):");

                String filePath = "files/exampleCSV.csv"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    // Создаем объект FileOutputStream для записи в файл
                    FileWriter fileWriter = new FileWriter(filePath);

                    // Записываем данные в файл
                    fileWriter.write(enteredText+"\n");

                    // Закрываем поток вывода файла
                    fileWriter.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        // Дописывание текста в файлы
        downBtnTXT[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");

                String filePath = "files/exampleTXT.txt"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    // Создаем объект FileOutputStream для записи в файл
                    FileOutputStream fileOutputStream = new FileOutputStream(filePath,true);

                    // Записываем данные в файл
                    byte[] byteData = enteredText.getBytes();
                    fileOutputStream.write(byteData);

                    // Закрываем поток вывода файла
                    fileOutputStream.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        downBtnDOCX[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");

                String filePath = "files/exampleDOCX.docx"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    FileInputStream fis = new FileInputStream(new File(filePath));
                    XWPFDocument document = new XWPFDocument(fis);
                    fis.close();

                    // Получаем последний параграф в документе
                    XWPFParagraph lastParagraph = document.getParagraphs().get(document.getParagraphs().size() - 1);

                    // Создаем новый параграф и добавляем его в документ
                    XWPFParagraph newParagraph = document.createParagraph();
                    newParagraph.createRun().setText(enteredText);

                    // Копируем текст из исходного параграфа в новый параграф
                    for (XWPFRun run : lastParagraph.getRuns()) {
                        String text = run.getText(0);
                        if (text != null) {
                            newParagraph.createRun().setText(text);
                        }
                    }

                    // Сохранение изменений в файл
                    FileOutputStream fos = new FileOutputStream(new File(filePath));
                    document.write(fos);
                    fos.close();

                    JOptionPane.showMessageDialog(null, "Данные успешно записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }
            }
        });

        downBtnXLS[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст который необходимо записать в файл:");
                String enteredText1 = JOptionPane.showInputDialog(null, "колонка:");
                String enteredText2 = JOptionPane.showInputDialog(null, "строка:");
                String filePath = "files/exampleXLS.xls"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    FileInputStream fis = new FileInputStream(new File(filePath));
                    HSSFWorkbook workbook = new HSSFWorkbook(fis); // Используем HSSFWorkbook для обработки файлов XLS
                    HSSFSheet sheet = workbook.getSheetAt(0);

                    // Запись данных в существующий файл XLS
                    HSSFRow row = sheet.createRow(Integer.parseInt(enteredText2) - 1);
                    HSSFCell cell = row.createCell(Integer.parseInt(enteredText1) - 1);
                    cell.setCellValue(enteredText);

                    // Сохранение изменений в файл
                    FileOutputStream fos = new FileOutputStream(new File(filePath));
                    workbook.write(fos);
                    fos.close();

                    JOptionPane.showMessageDialog(null, "Данные успешно записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }
            }
        });
        downBtnCSV[1].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String enteredText = JOptionPane.showInputDialog(null, "Введите текст в формате CSV(text,text,text):");

                String filePath = "files/exampleCSV.csv"; // Задайте путь и имя файла, который вы хотите открыть для записи

                try {
                    // Создаем объект FileOutputStream для записи в файл
                    FileWriter fileWriter = new FileWriter(filePath,true);

                    // Записываем данные в файл
                    fileWriter.write(enteredText);

                    // Закрываем поток вывода файла
                    fileWriter.close();

                    JOptionPane.showMessageDialog(null, "Данные успешено записаны!");
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }

            }
        });

        // Чтение файлов

        downBtnTXT[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                String line;
                String endInfoFromFile = "";

                String filePath = "files/exampleTXT.txt"; // Задайте путь и имя файла, который вы хотите открыть для записи
                try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {

                    while ((line = br.readLine()) != null) {
                        endInfoFromFile+=" \n "+line;
                    }
                    JOptionPane.showMessageDialog(null, "Данные успешено считаны: " + endInfoFromFile);

                } catch (IOException exception) {
                    System.out.println("Ошибка: " + exception.getMessage());
                }

            }
        });

        downBtnDOCX[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String filePath = "files/exampleDOCX.docx"; // Задайте путь и имя файла, который вы хотите открыть для записи
                String endInfoFromFile = "";
                try {
                    FileInputStream fis = new FileInputStream(new File(filePath));
                    XWPFDocument document = new XWPFDocument(fis);
                    fis.close();

                    // Перебираем все параграфы в документе
                    List<XWPFParagraph> paragraphs = document.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        String text = paragraph.getText();
                        endInfoFromFile+=" \n "+text;
                    }

                    JOptionPane.showMessageDialog(null, "Данные успешено считаны: " + endInfoFromFile);
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }
            }
        });

        downBtnXLS[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String filePath = "files/exampleXLS.xls"; // Задайте путь и имя файла, который вы хотите открыть для записи
                String endInfoFromFile = "\n";
                try (FileInputStream fis = new FileInputStream(filePath);
                     Workbook workbook = new HSSFWorkbook(fis)) {

                    Sheet sheet = workbook.getSheetAt(0); // Получаем первый лист в книге
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            CellType cellType = cell.getCellType();
                            if (cellType == CellType.STRING) {
                                endInfoFromFile+=cell.getStringCellValue() + "   ";
                            } else if (cellType == CellType.NUMERIC) {
                                endInfoFromFile+=cell.getNumericCellValue() + "   ";
                            } else if (cellType == CellType.BOOLEAN) {
                                endInfoFromFile+=cell.getBooleanCellValue() + "   ";
                            } else {
                                // Обработка других типов данных, если необходимо
                            }
                        }
                        endInfoFromFile+="\n";
                    }
                    JOptionPane.showMessageDialog(null, "Данные успешено считаны                      : " + endInfoFromFile);
                } catch (IOException exception) {
                    System.out.println("Error: " + exception.getMessage());
                }
            }
        });
        downBtnCSV[2].addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {

                String line;
                String endInfoFromFile = "";

                String filePath = "files/exampleCSV.csv"; // Задайте путь и имя файла, который вы хотите открыть для записи
                try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {

                    while ((line = br.readLine()) != null) {
                        endInfoFromFile+=" \n "+line;
                    }
                    JOptionPane.showMessageDialog(null, "Данные успешено считаны: " + endInfoFromFile);

                } catch (IOException exception) {
                    System.out.println("Ошибка: " + exception.getMessage());
                }


            }
        });



        setContentPane(panel);
        setSize(1280,720);
    }

    public void initVetical(JButton[] buttons, int coutnX, int coutnY,int step,boolean Visibale,int width,int height){
        for (int i=0;i<buttons.length;i++,coutnY+=step){
            buttons[i].setVisible(Visibale);
            buttons[i].setSize(width,height);
            buttons[i].setLocation(coutnX,coutnY);
            panel.add(buttons[i]);
        }
    }
    public void initHorizontal(JButton[] buttons, int coutnX, int coutnY,int step,boolean Visibale,int width,int height){
        for (int i=0;i<buttons.length;i++,coutnX+=step){
            buttons[i].setVisible(Visibale);
            buttons[i].setSize(width,height);
            buttons[i].setLocation(coutnX,coutnY);
            panel.add(buttons[i]);
        }

    }
    public static void editVisible(JButton[] buttons , boolean visibale){
        for (int i=0;i<buttons.length;i++){
            buttons[i].setVisible(visibale);
        }
    }
    public static void editlocation(JButton[] buttons,int countX,int countY,int step){
        for(int i=0;i<buttons.length;i++,countY+=step){
            buttons[i].setLocation(countX,countY);
        }
    }


}
