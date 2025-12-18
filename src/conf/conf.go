package conf

import (
	"fmt"

	// "github.com/spf13/cobra"
	"github.com/spf13/viper"
)

const INPUT = "input"
const HEADER_LINES = "header_lines"

//默认配置文件在用户目录
func ReadConfig() {

	// home, err:= os.UserHomeDir()

	// cobra.CheckErr(err)

	// viper.AddConfigPath(home)
	// viper.SetConfigName("config_common.yaml")

	viper.AddConfigPath("../conf")
	viper.SetConfigName("config.yaml")

	viper.SetConfigType("yaml")

	if err := viper.ReadInConfig(); err != nil {
		if _, ok := err.(viper.ConfigFileNotFoundError); ok {
			fmt.Println("not find " + err.Error())
		} else {
			fmt.Printf("read config file err, %v \n", err)
		}
	}
}
