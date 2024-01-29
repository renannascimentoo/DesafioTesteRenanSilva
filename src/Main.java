import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) throws Exception {

            // Carregar o arquivo Excel
            FileInputStream fis = new FileInputStream(new File("C:\\Users\\user\\Desktop\\DesafioTesteRenan\\src\\Cópia de Engenharia de Software - Desafio [RENAN DA SILVA].xlsx"));
            Workbook workbook = new XSSFWorkbook(fis);

            // Selecionar a planilha ativa
            Sheet sheet = workbook.getSheetAt(0);
            //cria uma lista de students//
            List<students> students = new LinkedList<>();
        // Itera sobre as linhas da planilha
        for (Row row : sheet) {
            double average = 0 ;
            // Pular a linha de cabeçalho
            if (row.getRowNum() < 3) {
                continue;
            }
            //atribuir valores//
            int registration = (int) row.getCell(0).getNumericCellValue();
            String name = row.getCell(1).getStringCellValue();
            int fouls = (int) row.getCell(2).getNumericCellValue();
            double p1 = row.getCell(3).getNumericCellValue();
            double p2 = row.getCell(4).getNumericCellValue();
            double p3 = row.getCell(5).getNumericCellValue();
            average = (p1 + p2 + p3) / 3;
            average = average / 10;
            // Determinar a situação do aluno
            double totalaulas = 60;
            String situation = null;
            double naf = 0;
            if (fouls > 60 * 0.25) {
                situation = "Reprovado por Falta";
                naf = 0;
            } else if (average < 5) {
                situation = "Reprovado por Nota";
                naf = 0;
            } else if (average < 7) {
                situation = "Exame Final";
                naf = (naf + average) / 2;
            } else if(average>7){
                situation = "Aprovado";
                naf = 0;
            }
            //adicionar na lista//
            students.add(new students(registration,name,fouls,p1,p2,p3,average,situation,naf));
            // Escrever o resultado na planilha
            Cell situationCell = row.createCell(6);
            situationCell.setCellValue(situation);

            Cell nafCell = row.createCell(7);
            nafCell.setCellValue(naf);
        }
        // Salvar o arquivo Excel
        FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\user\\Desktop\\DesafioTesteRenan\\src\\Cópia de Engenharia de Software - Desafio [RENAN DA SILVA].xlsx"));
        workbook.write(fos);
        workbook.close();
        fos.close();

        //mostrar na tela//
        for (students aluno: students) {
            System.out.println(aluno);
        }


    }
}
//classe students //
 class students{
    Integer registration;
    String name;
    int fouls;
    double p1;
    double p2;
    double p3;
    double average;
    String situation;
    double naf;
//construtor//

    public students(Integer registration, String name, int fouls, double p1, double p2, double p3, double average, String situation, double naf) {
        this.registration = registration;
        this.name = name;
        this.fouls = fouls;
        this.p1 = p1;
        this.p2 = p2;
        this.p3 = p3;
        this.average = average;
        this.situation = situation;
        this.naf = naf;
    }


    //getters e setters//


    public Integer getregistration() {
        return registration;
    }

    public void setregistration(Integer registration) {
        this.registration = registration;
    }

    public String getname() {
        return name;
    }

    public void setname(String name) {
        this.name = name;
    }

    public int getfouls() {
        return fouls;
    }

    public void setfouls(int fouls) {
        this.fouls = fouls;
    }

    public double getP1() {
        return p1;
    }

    public void setP1(double p1) {
        this.p1 = p1;
    }

    public double getP2() {
        return p2;
    }

    public void setP2(double p2) {
        this.p2 = p2;
    }

    public double getP3() {
        return p3;
    }

    public void setP3(double p3) {
        this.p3 = p3;
    }

    @Override
    public String toString() {
        return "students{" +
                "registration=" + registration +
                ", name='" + name + '\'' +
                ", fouls=" + fouls +
                ", p1=" + p1 +
                ", p2=" + p2 +
                ", p3=" + p3 +
                ", average=" + average +
                ", situation='" + situation + '\'' +
                ", naf=" + naf +
                '}';
    }
}