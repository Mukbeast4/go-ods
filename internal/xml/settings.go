package xml

type ConfigItem struct {
	Name  string
	Type  string
	Value string
}

type ConfigItemMapEntry struct {
	Name  string
	Items []ConfigItem
}

type SheetFreezeSettings struct {
	SheetName string
	FreezeCol int
	FreezeRow int
}
