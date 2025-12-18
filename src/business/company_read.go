package business

import (
	"electricity_bill/src/types"
	"electricity_bill/src/utils"
	"fmt"
	"log"
	"reflect"

	"github.com/spf13/viper"
	"github.com/tealeg/xlsx/v3"
)

const CONTENT_ROOM_NO = "门牌号"
const CONTENT_COMPANY = "单位名称"
const CONTENT_CONTACT = "联系人"
const CONTENT_PHONE = "电话"
const CONTENT_NEED_BILL = "需要账单"

func ReadCompany(c *chan (types.CompanyInfo), cFinish *chan (string)) error {
	file, err := xlsx.OpenFile(viper.GetString("elec_file"))

	if err != nil {
		panic(err)
		// return err
	}

	// readSheetSingleCompany(file.Sheets[0], &companies)
	if sheet, found := file.Sheet[viper.GetString("company_sheet")]; found {

		err = readSheetCompanyInfo(sheet, c)

		if err != nil {
			panic(err)
		}
	} else {
		panic("can not get company info !!!")
	}

	*cFinish <- "company finish"
	return nil
}

func readSheetCompanyInfo(sheet *xlsx.Sheet, c *chan (types.CompanyInfo)) error {
	headers := make([]string, sheet.MaxCol)
	err := readCompanyHeaders(sheet, &headers)
	if err != nil {
		return err
	}

	err = readCompanyData(sheet, &headers, c)

	fmt.Printf("-------- \n %+v \n", c)

	if err != nil {
		return err
	}
	return nil
}

func readCompanyHeaders(sheet *xlsx.Sheet, headers *[]string) error {
	headerLines := viper.GetInt("company_header_lines")

	for rowIndex := range headerLines {
		row, err := sheet.Row(rowIndex)
		if err != nil {
			fmt.Println("read company err :", err.Error())
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
	}
	return nil
}

func readCompanyData(sheet *xlsx.Sheet, headers *[]string, c *chan (types.CompanyInfo)) error {
	headerLines := viper.GetInt("company_header_lines")

	for rowIndex := headerLines; rowIndex < sheet.MaxRow; rowIndex++ {
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
				cip.IsAddPay = v == "1-1-总"

			case "IsNeedBill":
				cip.IsNeedBill = cell.Value == "是"
			default:
				reflect.ValueOf(&cip).Elem().FieldByName((*headers)[colIndex]).SetString(cell.Value)
			}
		}
		*c <- cip
	}
	return nil
}
