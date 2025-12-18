package utils

import (
	"strconv"
	"strings"
	"time"
)

func FindUnitFromRoomNo(roomNo string, unit *int, floor *int) error {
	s := strings.Split(roomNo, "-")
	u, err := strconv.Atoi(s[0])
	if err == nil {
		*unit = u
	} else {
		return err
	}

	f, err := strconv.Atoi(s[1][0:1])
	if err == nil {
		*floor = f
	} else {
		return err
	}

	return nil
}

func DaysInMonth(year int, month int) int {
	t := time.Date(year, time.Month(month+1), 1, 0, 0, 0, 0, time.UTC)
	t = t.AddDate(0, 0, -1)
	return t.Day()
}
