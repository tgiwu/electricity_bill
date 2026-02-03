package business

import (
	"electricity_bill/src/types"
	"sort"

	"github.com/spf13/viper"
)

func ConstructNotiItems(companies *map[int]map[string]types.CompanyInfo, indics *map[string]types.Indication, notiMap *map[int][]types.NotificationItem) {

	for unit, infoMap := range *companies {
		list := []types.NotificationItem{}
		constructNotiItemsUnit(&infoMap, indics, &list)

		if len(list) > 0 {
			(*notiMap)[unit] = list
		}
	}
}

func constructNotiItemsUnit(infoMap *map[string]types.CompanyInfo, indics *map[string]types.Indication, notiList *[]types.NotificationItem) {
	keys := sortByGateNo(infoMap)
	for _, key := range keys {
		if !(*infoMap)[key].IsNeedBill {
			continue
		}
		item := types.NotificationItem{CompanyInfo: (*infoMap)[key],
			Liquidated:   viper.GetFloat64("liquidated"),
			IndicList:    &[]types.TableRow{},
			AirIndicList: &[]types.TableRow{}}

		constructSingleItem(indics, &item)
		*notiList = append(*notiList, item)
	}
}

func constructSingleItem(indics *map[string]types.Indication, item *types.NotificationItem) {
	for _, gateNo := range item.GateNos {

		tableRow := types.TableRow{
			RoomNo:         gateNo,
			Month:          viper.GetInt("target_month"),
			IndicLastMonth: (*indics)[gateNo].IndicLastMonth,
			IndicCurrent:   (*indics)[gateNo].Indic,
			Times:          (*indics)[gateNo].Times,
			Cost:           (*indics)[gateNo].Cost,
			Payment:        (*indics)[gateNo].Cost * item.RateOfPay,
		}
		item.CostSum += (*indics)[gateNo].CostAll

		*item.IndicList = append(*item.IndicList, tableRow)

		if (*indics)[gateNo].CostAirControal != 0 {
			tableRow := types.TableRow{
				RoomNo:  gateNo,
				Month:   viper.GetInt("target_month"),
				Cost:    (*indics)[gateNo].CostAirControal,
				Payment: (*indics)[gateNo].CostAirControal * item.RateOfPay,
			}
			*item.AirIndicList = append(*item.AirIndicList, tableRow)
		}

		item.Payment += (*indics)[gateNo].Cost * viper.GetFloat64("price")
		item.PaymentAll += (*indics)[gateNo].CostAll * viper.GetFloat64("price")
	}

}

func sortByGateNo(companies *map[string]types.CompanyInfo) []string {

	keys := []string{}
	for key := range *companies {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	return keys
}
