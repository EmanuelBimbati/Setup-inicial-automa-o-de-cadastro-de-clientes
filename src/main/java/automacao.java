import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;

public class automacao {
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\bimba\\Documents\\driverchorme\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        try {

            driver.get("http://digipecas01.maxdatacenter.com.br/manelaoautopecas/Default.aspx");

            wait.until(ExpectedConditions.presenceOfElementLocated(By.id("Login1_LoginButton")));


            driver.findElement(By.id("Login1_UserName")).sendKeys("manelaoautopecas");
            driver.findElement(By.id("Login1_Password")).sendKeys("123");
            driver.findElement(By.id("Login1_LoginButton")).click();

            Thread.sleep(1000); // Aguarda redirecionamento


            FileInputStream file = new FileInputStream(new File("C:\\Users\\bimba\\Documents\\digipecas-20250716T215959Z-1-001\\clientes.xlsx"));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);


            Row ultimaLinha = null;
            int lastRowNum = sheet.getLastRowNum();

            for (int i = lastRowNum; i >= 0; i--) {
                Row row = sheet.getRow(i);
                if (row != null && row.getCell(0) != null && row.getCell(0).getCellType() != CellType.BLANK) {
                    ultimaLinha = row;
                    break;
                }
            }

            if (ultimaLinha != null) {
                try {

                    String nome = ultimaLinha.getCell(0).getStringCellValue();

                    Cell cellCpf = ultimaLinha.getCell(1);
                    String cpf = (cellCpf != null && cellCpf.getCellType() == CellType.NUMERIC)
                            ? String.valueOf((long) cellCpf.getNumericCellValue())
                            : (cellCpf != null ? cellCpf.getStringCellValue() : "");

                    String tipo = ultimaLinha.getCell(2) != null ? ultimaLinha.getCell(2).getStringCellValue() : "";

                    String email;
                    Cell cellEmail = ultimaLinha.getCell(3);
                    if (cellEmail == null) {
                        email = "naotem@gmail.com";
                    } else if (cellEmail.getCellType() == CellType.STRING) {
                        String valor = cellEmail.getStringCellValue().trim();
                        email = valor.isEmpty() ? "naotem@gmail.com" : valor;
                    } else if (cellEmail.getCellType() == CellType.NUMERIC) {
                        email = String.valueOf((long) cellEmail.getNumericCellValue());
                    } else {
                        email = "naotem@gmail.com";
                    }


                    Cell cellNum = ultimaLinha.getCell(4);
                    String num = (cellNum != null && cellNum.getCellType() == CellType.NUMERIC)
                            ? String.valueOf((long) cellNum.getNumericCellValue())
                            : (cellNum != null ? cellNum.getStringCellValue() : "");

                    Cell cellCep = ultimaLinha.getCell(5);
                    String cep = (cellCep != null && cellCep.getCellType() == CellType.NUMERIC)
                            ? String.valueOf((long) cellCep.getNumericCellValue())
                            : (cellCep != null ? cellCep.getStringCellValue() : "");

                    Cell cellNumCasa = ultimaLinha.getCell(6);
                    String casa = (cellNumCasa != null && cellNumCasa.getCellType() == CellType.NUMERIC)
                            ? String.valueOf((long) cellNumCasa.getNumericCellValue())
                            : (cellNumCasa != null ? cellNumCasa.getStringCellValue() : "");


                    driver.get("http://digipecas01.maxdatacenter.com.br/manelaoautopecas/cadastros/pessoas_form.aspx");

                    wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ctl00_ContentPlaceHolder1_dvPessoa_TextBox3")));


                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_TextBox3")).sendKeys(nome);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_TextBox2")).sendKeys(nome);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_txtCNPJ")).sendKeys(cpf);

                    if (tipo.equalsIgnoreCase("fisica")) {
                        driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_RadioButtonList1_1")).click();
                    } else {
                        driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_RadioButtonList1_0")).click();
                    }

                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_TextBox9")).sendKeys(email);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_edtTelefone")).sendKeys(num);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_edtCep")).sendKeys(cep);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_btnConsultaCEP")).click();

                    Thread.sleep(500); //

                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_TextBox6")).sendKeys(casa);
                    driver.findElement(By.id("ctl00_ContentPlaceHolder1_dvPessoa_ucCommandInsert1_btnSalvar")).click();

                    Thread.sleep(1000);

                } catch (Exception linhaErro) {
                    System.out.println("Erro ao cadastrar cliente: " + linhaErro.getMessage());
                }
            }

            workbook.close();
            file.close();

        } catch (Exception e) {
            e.printStackTrace();
        }  {

        }
    }
}
