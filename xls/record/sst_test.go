package record

import "testing"

func TestSSTContinueWideToCompressed(t *testing.T) {
	// Seven characters in total. The first record contains "Я1" in UTF-16LE.
	sst := SST{RgbSrc: []byte{7, 0, 1, 0x2F, 0x04, '1', 0}}
	sst.Read("", 0, 0)

	// The continuation contains ".2025" in compressed form,
	// followed by a separate one-character string, "X".
	sst.RgbSrc = append(sst.RgbSrc, []byte(".2025")...)
	sst.RgbSrc = append(sst.RgbSrc, 1, 0, 0, 'X')
	sst.Read("continue", 0, 7)

	if len(sst.Rgb) != 2 {
		t.Fatalf("got %d shared strings, want 2", len(sst.Rgb))
	}
	for i, want := range []string{"Я1.2025", "X"} {
		if got := sst.Rgb[i].String(); got != want {
			t.Errorf("string %d: got %q, want %q", i, got, want)
		}
	}
}
