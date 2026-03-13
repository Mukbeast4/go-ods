package goods

import "testing"

func TestCellNameToCoordinates(t *testing.T) {
	tests := []struct {
		cell    string
		col     int
		row     int
		wantErr bool
	}{
		{"A1", 1, 1, false},
		{"B3", 2, 3, false},
		{"Z1", 26, 1, false},
		{"AA1", 27, 1, false},
		{"AZ1", 52, 1, false},
		{"BA1", 53, 1, false},
		{"a1", 1, 1, false},
		{"AB10", 28, 10, false},
		{"", 0, 0, true},
		{"A", 0, 0, true},
		{"1", 0, 0, true},
		{"A0", 0, 0, true},
		{"1A", 0, 0, true},
	}

	for _, tt := range tests {
		t.Run(tt.cell, func(t *testing.T) {
			col, row, err := CellNameToCoordinates(tt.cell)
			if (err != nil) != tt.wantErr {
				t.Errorf("CellNameToCoordinates(%q) error = %v, wantErr %v", tt.cell, err, tt.wantErr)
				return
			}
			if col != tt.col || row != tt.row {
				t.Errorf("CellNameToCoordinates(%q) = (%d, %d), want (%d, %d)", tt.cell, col, row, tt.col, tt.row)
			}
		})
	}
}

func TestCoordinatesToCellName(t *testing.T) {
	tests := []struct {
		col     int
		row     int
		want    string
		wantErr bool
	}{
		{1, 1, "A1", false},
		{2, 3, "B3", false},
		{26, 1, "Z1", false},
		{27, 1, "AA1", false},
		{52, 1, "AZ1", false},
		{53, 1, "BA1", false},
		{28, 10, "AB10", false},
		{0, 1, "", true},
		{1, 0, "", true},
		{-1, 1, "", true},
	}

	for _, tt := range tests {
		t.Run(tt.want, func(t *testing.T) {
			got, err := CoordinatesToCellName(tt.col, tt.row)
			if (err != nil) != tt.wantErr {
				t.Errorf("CoordinatesToCellName(%d, %d) error = %v, wantErr %v", tt.col, tt.row, err, tt.wantErr)
				return
			}
			if got != tt.want {
				t.Errorf("CoordinatesToCellName(%d, %d) = %q, want %q", tt.col, tt.row, got, tt.want)
			}
		})
	}
}

func TestRoundTrip(t *testing.T) {
	for col := 1; col <= 1000; col++ {
		name := columnNumberToName(col)
		got := columnNameToNumber(name)
		if got != col {
			t.Errorf("round-trip failed for col %d: name=%q, back=%d", col, name, got)
		}
	}
}
