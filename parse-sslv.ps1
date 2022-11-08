Import-Module "C:\selenium\WebDriver.dll"

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
$ChromeDriver.Navigate().GoToUrl("https://www.ss.lv/ru/transport/cars/")

$carArr_1stCol = [System.Collections.ArrayList]::new()
$carArr_2stCol = [System.Collections.ArrayList]::new()

#collect all car names in first column and add to arraylist
for ($i=1; $i -cle 14; $i++){
    $carName = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[4]/div/table/tbody/tr/td/div[1]/table/tbody/tr/td/form/table[2]/tbody/tr/td[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td/h4[$i]/a")).GetAttribute("text")
    $carHref = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[4]/div/table/tbody/tr/td/div[1]/table/tbody/tr/td/form/table[2]/tbody/tr/td[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td/h4[$i]/a")).GetAttribute("href")
    $carArr_1stCol.Add($carName)
    $carArr_2stCol.Add($carHref)

}

#go to link and open a car
for ($i=1;$i -cle 1; $i++){
    $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[4]/div/table/tbody/tr/td/div[1]/table/tbody/tr/td/form/table[2]/tbody/tr/td[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td/h4[$i]/a")).Click()
    $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[4]/div/table/tbody/tr/td/div[1]/table/tbody/tr/td/form/table[1]/tbody/tr/td[1]/table/tbody/tr/td[1]/span/input[1]")).SendKeys("1000")
    $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[4]/div/table/tbody/tr/td/div[1]/table/tbody/tr/td/form/table[1]/tbody/tr/td[1]/table/tbody/tr/td[1]/span/input[2]")).SendKeys("5000")
}
#add criterias

#open car

#get date and go out




#$ExcelObj = New-Object -comobject Excel.Application
#$ExcelObj.Visible = $true

#$ExcelWorkBook = $ExcelObj.Workbooks.Add()
#$ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)
# Rename a worksheet
#$ExcelWorkSheet.Name = 'Cars'

# Fill in the head of the table
#$ExcelWorkSheet.Cells.Item(1,1) = 'Car model'
#$ExcelWorkSheet.Cells.Item(1,2) = 'Link'

#for (($i=2),($j=0);$i -cle 15; $i++,$j++){
#    $ExcelWorkSheet.Columns.Item(1).Rows.Item($i) = $carArr_1stCol[$j]
#    $ExcelWorkSheet.Columns.Item(2).Rows.Item($i) = $carArr_2stCol[$j]
#}

# Make the table head bold, set the font size and the column width
#$ExcelWorkSheet.Rows.Item(1).Font.Bold = $true
#$ExcelWorkSheet.Rows.Item(1).Font.size=15

# Save the report and close Excel:
#$ExcelWorkBook.SaveAs('D:\Desktop\sslv-report.xlsx')
#$ExcelWorkBook.close($true)