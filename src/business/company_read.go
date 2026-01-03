package business

import (
	"electricity_bill/src/types"
	"electricity_bill/src/utils"
	"encoding/json"
	"fmt"
	"log"
	"reflect"
	"strings"

	"github.com/spf13/viper"
	"github.com/tealeg/xlsx/v3"
)

const CONTENT_ROOM_NO = "门牌号"
const CONTENT_COMPANY = "单位名称"
const CONTENT_CONTACT = "联系人"
const CONTENT_PHONE = "电话"
const CONTENT_NEED_BILL = "账单"

var dataRowStart = 0

func ReadCompany(c *chan (types.CompanyInfo), cFinish *chan (string)) {
	file, err := xlsx.OpenFile(viper.GetString("elec_file"))

	if err != nil {
		log.Panic(err)
		// return err
	}

	if sheet, found := file.Sheet[viper.GetString("company_sheet")]; found {

		err = readSheetCompanyInfo(sheet, c)

		if err != nil {
			log.Panic(err)
		}
	} else {
		log.Panic("can not get company info !!!")
	}

	*cFinish <- "com_f"
}

func readSheetCompanyInfo(sheet *xlsx.Sheet, c *chan (types.CompanyInfo)) error {
	headers := make([]string, sheet.MaxCol)
	err := readCompanyHeaders(sheet, &headers)
	if err != nil {
		return err
	}

	err = readCompanyData(sheet, &headers, c)

	if err != nil {
		return err
	}
	return nil
}

func readCompanyHeaders(sheet *xlsx.Sheet, headers *[]string) error {

	for rowIndex := 0; rowIndex < sheet.MaxRow; rowIndex++ {
		row, err := sheet.Row(rowIndex)
		if err != nil {
			fmt.Println("read company err :", err.Error())
			continue
		}

		if row.GetCell(0).Value != "门牌号" {
			continue
		}

		for colIndex := 0; colIndex < sheet.MaxCol; colIndex++ {
			v := row.GetCell(colIndex).Value

			switch v {
			case CONTENT_ROOM_NO:
				(*headers)[colIndex] = "GateNo"
			case CONTENT_COMPANY:
				(*headers)[colIndex] = "Name"
			case CONTENT_CONTACT:
				(*headers)[colIndex] = "Contact"
			case CONTENT_PHONE:
				(*headers)[colIndex] = "Phone"
			case CONTENT_NEED_BILL:
				(*headers)[colIndex] = "IsNeedBill"
			default:
				//ignore
			}
		}
		dataRowStart = rowIndex + 1
	}
	return nil
}

func readCompanyData(sheet *xlsx.Sheet, headers *[]string, c *chan (types.CompanyInfo)) error {

	for rowIndex := dataRowStart; rowIndex < sheet.MaxRow; rowIndex++ {
		row, err := sheet.Row(rowIndex)
		if err != nil {
			return err
		}

		cip := types.CompanyInfo{IsNeedBill: true}
		for colIndex := 0; colIndex < sheet.MaxCol; colIndex++ {
			cell := row.GetCell(colIndex)

			switch (*headers)[colIndex] {
			case "GateNo":
				v := cell.Value

				err = utils.FindUnitFromRoomNo(v, &cip.Unit, &cip.Floor)
				if err != nil {
					log.Panic(err)
				}
				cip.GateNo = v

			case "IsNeedBill":
				cip.IsNeedBill = cell.Value == "是"
			default:
				reflect.ValueOf(&cip).Elem().FieldByName((*headers)[colIndex]).SetString(cell.Value)
			}
		}
		if strings.Contains(cip.Name, ";") {
			names := strings.Split(cip.Name, ";")
			for _, n := range names {
				copied := types.CompanyInfo{}
				jsonData, err := json.Marshal(cip)
				if err != nil {
					log.Fatal(err)
				}
				err = json.Unmarshal(jsonData, &copied)
				if err != nil {
					log.Fatal(err)
				}
				copied.Name = n
				copied.IsAddPayment = true
				copied.IsNeedBill = true
				copied.LookUpKey = fmt.Sprintf("%s@%s", copied.GateNo, copied.Name)
				copied.RateOfPay = 1.0 / float64(len(names))

				*c <- copied
			}
		} else {
			cip.LookUpKey = fmt.Sprintf("%s@%s", cip.GateNo, cip.Name)
			*c <- cip
		}
	}
	return nil
}
