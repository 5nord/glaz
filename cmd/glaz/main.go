package main

import (
	"fmt"
	"os"
	"os/exec"
	"time"

	"github.com/5nord/glaz"
	"github.com/spf13/cobra"
)

var (
	rootCmd = &cobra.Command{
		Use:  "glaz",
		RunE: state,
	}

	inCmd = &cobra.Command{
		Use:  "in",
		RunE: in,
	}

	outCmd = &cobra.Command{
		Use:  "out",
		RunE: out,
	}

	openCmd = &cobra.Command{
		Use:  "open",
		RunE: open,
	}

	file = os.ExpandEnv(fmt.Sprintf("$HOME/Documents/Arbeit/GLAZ/GLAZFormular_%d_deutsch.xlsx", time.Now().Year()))
)

func main() {
	rootCmd.AddCommand(inCmd, outCmd, openCmd)
	rootCmd.Execute()
}

func open(cmd *cobra.Command, args []string) error {
	proc := exec.Command("xdg-open", file)
	proc.Start()
	return nil
}

func state(cmd *cobra.Command, args []string) error {
	fmt.Printf("Opening %s\n", file)
	sheet, err := glaz.OpenFile(file)
	if err == nil {
		fmt.Printf("[32;1m+ %s[0m\n", sheet.Today())
	}
	return err
}

func in(cmd *cobra.Command, args []string) error {
	fmt.Printf("Opening %s\n", file)
	sheet, err := glaz.OpenFile(file)
	if err != nil {
		return err
	}

	now := time.Now()
	day := sheet.Today()

	fmt.Printf("[31;1m- %s[0m\n", day)

	if day.Work.Begin.IsZero() || day.Work.Begin.After(now) {
		day.Work.Begin = now
	}

	if !day.Work.End.IsZero() && day.Work.End.Before(now) {
		pause := now.Sub(day.Work.End)
		day.Pause.End = day.Pause.End.Add(pause)
		day.Work.End = now
	}

	if err := sheet.Err(); err != nil {
		return err
	}

	fmt.Printf("[32;1m+ %s[0m\n", day)
	return sheet.Update(day)
}

func out(cmd *cobra.Command, args []string) error {
	sheet, err := glaz.OpenFile(file)
	if err != nil {
		return err
	}

	now := time.Now()
	day := sheet.Today()

	fmt.Printf("[31;1m- %s[0m\n", day)

	if day.Work.End.IsZero() || day.Work.End.Before(now) {
		day.Work.End = now
	}

	if err := sheet.Err(); err != nil {
		return err
	}

	fmt.Printf("[32;1m+ %s[0m\n", day)
	return sheet.Update(day)
}
