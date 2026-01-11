package business

import (
	"electricity_bill/src/types"
	"sort"
)

func ConstructNotiItems(companies *map[int]map[string]types.CompanyInfo, indics *map[string]types.Indication, notiMap *map[int][]types.NotificationItem) {

	for unit, infoMap := range *companies {
		list := []types.NotificationItem{}
		constructNotiItemsUnit(&infoMap, indics, unit, &list)

		if len(list) > 0 {
			(*notiMap)[unit] = list
		}
	}
}

func constructNotiItemsUnit(infoMap *map[string]types.CompanyInfo, indics *map[string]types.Indication, unit int, notiList *[]types.NotificationItem) {
	keys := sortByGateNo(infoMap)
	for _, key := range keys {
		item := types.NotificationItem{CompanyInfo: (*infoMap)[key],
			Liquidated:   0.00,
			IndicList:    &[]types.TableRow{},
			AirIndicList: &[]types.TableRow{}}

		info := (*infoMap)[key]
		constructSingleItem(&info, indics, &item)
		*notiList = append(*notiList, item)
	}
}

func constructSingleItem(info *types.CompanyInfo, indics *map[string]types.Indication, item *types.NotificationItem) {
	
}

func sortByGateNo(companies *map[string]types.CompanyInfo) []string {

	keys := []string{}
	for key := range *companies {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	return keys
}
