package types

import "fmt"

// 区域
type Area struct {
	Name     string  //名称
	Previous float64 //上月度数
	Times    float64 //倍率
	Current  float64 //本月度数
	Kwh      float64 //本月消耗度数
}

// 计算耗能
func (a *Area) ClacKwh() {
	if a.Times == 0 {
		a.Times = 1
	}
	a.Kwh = (a.Current - a.Previous) * a.Times
}

// 每层每月记录
type AreasByMonth struct {
	NO         int
	Month      int
	All        Area //总电量
	East       Area //东区
	West       Area //西区
	Public     Area //公区
	Emergency  Area //应急
	AirControl Area //空调
}

// 楼层
type Floor struct {
	NO    int            //楼层
	Labms []AreasByMonth //每月各区域数据
}

// 耗电量
type Indication struct {
	RoomNo string  //房间号
	Times  float64 //倍率
	Unit   int     //单元
	Floor  int     //楼层
	// Original        float64 //表底
	IndicLastMonth  float64 //上月读数
	Indic           float64 //本月读数
	CostAirControal float64 //空调耗电量
	Cost            float64 //耗电量（不包含空调）
	CostAll         float64 //总耗电量
}

// 公司信息
type CompanyInfo struct {
	Name   string //公司名称
	GateNo string //门牌号
	// AreaName   string //区域名称
	Unit       int    //单元
	Floor      int    //楼层
	Contact    string //联系人
	Phone      string //电话
	IsNeedBill bool   //是否需要账单，默认只要由公司信息就出账单
	IsAddPay   bool   //是否添加应缴
}

// myError
type MyError struct {
	Path string
	Op   string
}

func (p MyError) Error() string {
	return fmt.Sprintf("%s %s ", p.Path, p.Op)
}
