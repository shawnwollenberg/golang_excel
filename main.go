package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func cTime(t string) string {
	//t := time.Now()
	sHold := strings.Split(t, "/")
	return (sHold[2] + "_" + ("0" + sHold[1])[0:2] + "_" + ("0" + sHold[0])[0:2])
}
func excelDate(excelNum int) string {
	startDate := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)
	finalDt := startDate.AddDate(0, 0, excelNum)
	return finalDt.Format("2006-01-02")
}

func finalTime(t string) string {
	sHold := strings.Split(t, "/")
	sDay := "0" + sHold[1]
	sMo := "0" + sHold[0]
	return sDay[len(sDay)-2:] + "/" + sMo[len(sMo)-2:] + "/" + sHold[2]
}
func getTime() string {
	t := time.Now()
	return t.Format("2006-01-02-15-04-05")
}
func indexOf(element string, data []string) int {
	for k, v := range data {
		if element == v {
			return k
		}
	}
	return -1 //not found.
}
func stringInSlice(a string, list []string) bool {
	for _, b := range list {
		if b == a {
			return true
		}
	}
	return false
}

type itemInfo struct {
	CurTime    string
	QuoteDt    string
	SE         string
	QuoteNum   string
	OppNum     string
	CustName   string
	TypeClient string
	SatClient  string
	Profil     string
	Concurrent string
	Avant      string
	Reason     string
	Qty        string
	Item       string
	RefPrice   string
	QuotePrice string
}

type DB []itemInfo

var db = DB{}

func Create(s itemInfo) {
	db = append(db, s)
	return
}
func main() {
	searchDir := "C:\\Test"
	files, err := ioutil.ReadDir(searchDir)
	if err != nil {
		log.Fatal(err)
	}
	//sort.Sort(ByModTime(files))

	for _, file := range files {
		if !file.IsDir() {
			xCurTime := getTime()
			sOldFileName := file.Name()
			xlFile, err := xlsx.OpenFile(searchDir + "\\" + sOldFileName)
			if err != nil {
				panic(err)
			}
			var sheet *xlsx.Sheet
			//sheet = xlFile.Sheets[0]
			/*ictr := 0
			for _, sheet := range xlFile.Sheets {
				fmt.Println(ictr, sheet.Name)
				ictr++
			}*/
			sheet = xlFile.Sheet["Civilité"]
			sQuoteDt := sheet.Rows[1].Cells[1].Value //B2
			sSE := sheet.Rows[2].Cells[1].Value      //B3
			sheet = xlFile.Sheet["Demande de remise"]
			sQuoteNum := sheet.Rows[5].Cells[1].Value
			sOppNum := sheet.Rows[3].Cells[1].Value
			sCustName := sheet.Rows[6].Cells[1].Value
			sTypeClient := sheet.Rows[17].Cells[1].Value
			sSatisfactionClient := sheet.Rows[18].Cells[1].Value
			sProfil := sheet.Rows[19].Cells[1].Value
			sConcurrent := sheet.Rows[20].Cells[1].Value
			sAvant := sheet.Rows[21].Cells[1].Value
			sReason := sheet.Rows[12].Cells[2].Value
			sheet = xlFile.Sheet["offre_de_prix"]
			for i := 54; i < 6074; i++ {
				sQtyStr := sheet.Rows[i].Cells[7].Value
				sQtyNum, _ := strconv.Atoi(sQtyStr) //H
				if sQtyNum > 0 {
					sItem := sheet.Rows[i].Cells[0].Value       //A
					sRefPrice := sheet.Rows[i].Cells[8].Value   //I
					sQuotePrice := sheet.Rows[i].Cells[9].Value //J

					s1 := itemInfo{xCurTime, sQuoteDt, sSE, sQuoteNum, sOppNum, sCustName, sTypeClient, sSatisfactionClient, sProfil, sConcurrent, sAvant, sReason, sQtyStr, sItem, sRefPrice, sQuotePrice}
					Create(s1)
				}
			}
		}
	}

	var file *xlsx.File
	var row *xlsx.Row
	var cell *xlsx.Cell

	xlsx.SetDefaultFont(8, "Arial")
	file = xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")
	row = sheet.AddRow()
	cell = row.AddCell()
	cell.Value = "Data Capture Date"
	cell = row.AddCell()
	cell.Value = "Quote Date"
	cell = row.AddCell()
	cell.Value = "SE"
	cell = row.AddCell()
	cell.Value = "Quote Number"
	cell = row.AddCell()
	cell.Value = "Opportunity Number"
	cell = row.AddCell()
	cell.Value = "Customer Name"
	cell = row.AddCell()
	cell.Value = "Item"
	cell = row.AddCell()
	cell.Value = "Quantity"
	cell = row.AddCell()
	cell.Value = "Reference Price"
	cell = row.AddCell()
	cell.Value = "Quoted Price"
	cell = row.AddCell()
	cell.Value = "Type Client"
	cell = row.AddCell()
	cell.Value = "Satisfaction Client"
	cell = row.AddCell()
	cell.Value = "Profil acheteur"
	cell = row.AddCell()
	cell.Value = "Concurrent"
	cell = row.AddCell()
	cell.Value = "Avantage produit"
	cell = row.AddCell()
	cell.Value = "Reason for discount"
	for i := 0; i < len(db); i++ {
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = db[i].CurTime
		cell = row.AddCell()
		holdDt1, _ := strconv.Atoi(db[i].QuoteDt)
		holdDt := excelDate(holdDt1)
		cell.Value = holdDt
		cell = row.AddCell()
		cell.Value = db[i].SE
		cell = row.AddCell()
		cell.Value = db[i].QuoteNum
		cell = row.AddCell()
		cell.Value = db[i].OppNum
		cell = row.AddCell()
		cell.Value = db[i].CustName
		cell = row.AddCell()
		cell.Value = db[i].Item
		cell = row.AddCell()
		cell.Value = db[i].Qty
		cell = row.AddCell()
		cell.Value = db[i].RefPrice
		cell = row.AddCell()
		cell.Value = db[i].QuotePrice
		cell = row.AddCell()
		cell.Value = db[i].TypeClient
		cell = row.AddCell()
		cell.Value = db[i].SatClient
		cell = row.AddCell()
		cell.Value = db[i].Profil
		cell = row.AddCell()
		cell.Value = db[i].Concurrent
		cell = row.AddCell()
		cell.Value = db[i].Avant
		cell = row.AddCell()
		cell.Value = db[i].Reason
	}
	err = file.Save("DataCollection_" + getTime() + ".xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}

}
