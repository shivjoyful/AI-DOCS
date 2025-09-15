function Format-CMMI7SOP {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    $word = $null
    $doc = $null

    try {
        # Create Word Application
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # Suppress alerts

        # Open Document with Error Handling
        try {
            $doc = $word.Documents.Open($InputPath, $false, $false)
            if (-not $doc) {
                throw "Could not open the document. Ensure the file is not corrupt or in use."
            }
        } catch {
            throw "Document Open Error: $_"
        }

        # Safe Page Setup
        try {
            $pageSetup = $doc.PageSetup
            $pageSetup.TopMargin = 72     # 1 inch
            $pageSetup.BottomMargin = 72  # 1 inch
            $pageSetup.LeftMargin = 90    # 1.25 inches
            $pageSetup.RightMargin = 90   # 1.25 inches
        } catch {
            Write-Warning "Could not set page margins: $_"
        }

        # Header and Footer Handling
        try {
            # Ensure at least one section exists
            if ($doc.Sections.Count -eq 0) {
                Write-Warning "No sections found in the document"
            } else {
                $section = $doc.Sections.Item(1)

                # Header
                $header = $section.Headers.Item(1)
                $header.Range.Text = "CONFIDENTIAL - [Company Name] | CMMI Level 7 Optimizing Process"
                $header.Range.Font.Name = "Calibri"
                $header.Range.Font.Size = 9
                $header.Range.Font.Color = 8421504  # Dark Gray
                $header.Range.ParagraphFormat.Alignment = 2  # Right aligned

                # Footer
                $footer = $section.Footers.Item(1)
                $footer.Range.Text = "Document Control: [Document ID] | Version: 1.0 | Last Updated: $(Get-Date -Format 'yyyy-MM-dd')"
                $footer.Range.Font.Name = "Calibri"
                $footer.Range.Font.Size = 8
                $footer.Range.Font.Color = 8421504  # Dark Gray
                $footer.Range.ParagraphFormat.Alignment = 2  # Right aligned
            }
        } catch {
            Write-Warning "Could not format header/footer: $_"
        }

        # Global Font Settings
        try {
            $doc.Content.Font.Name = "Arial"
            $doc.Content.Font.Size = 11
        } catch {
            Write-Warning "Could not set global font: $_"
        }

        # Custom Styles Creation with Error Handling
        $styles = @{
            "SOP Title" = @{
                FontName = "Calibri"
                FontSize = 18
                Bold = $true
                Color = 0
                SpaceBefore = 0
                SpaceAfter = 12
                Alignment = 1
            }
            "Section Heading" = @{
                FontName = "Calibri"
                FontSize = 14
                Bold = $true
                Color = 0
                SpaceBefore = 12
                SpaceAfter = 6
                Alignment = 0
            }
            "Subsection Heading" = @{
                FontName = "Calibri"
                FontSize = 12
                Bold = $true
                Color = 39424
                SpaceBefore = 10
                SpaceAfter = 4
                Alignment = 0
            }
        }

        # Create Styles Safely
        foreach ($styleName in $styles.Keys) {
            try {
                # Check if style already exists
                $existingStyle = $null
                try {
                    $existingStyle = $doc.Styles.Item($styleName)
                } catch {
                    # Style doesn't exist, create new
                }

                if (-not $existingStyle) {
                    $newStyle = $doc.Styles.Add($styleName, 1)
                } else {
                    $newStyle = $existingStyle
                }

                $styleConfig = $styles[$styleName]

                $newStyle.Font.Name = $styleConfig.FontName
                $newStyle.Font.Size = $styleConfig.FontSize
                $newStyle.Font.Bold = $styleConfig.Bold
                $newStyle.Font.Color = $styleConfig.Color

                $newStyle.ParagraphFormat.SpaceBefore = $styleConfig.SpaceBefore
                $newStyle.ParagraphFormat.SpaceAfter = $styleConfig.SpaceAfter
                $newStyle.ParagraphFormat.Alignment = $styleConfig.Alignment
            } catch {
                Write-Warning "Could not create style $styleName : $_"
            }
        }

        # Paragraph Formatting with Robust Error Handling
        for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
            try {
                $para = $doc.Paragraphs.Item($i)
                $text = $para.Range.Text.Trim()

                # Apply Specialized Formatting
                switch -Regex ($text) {
                    '^(Objective|Purpose):' { 
                        try { $para.Style = "Section Heading" } catch { Write-Warning "Could not set Section Heading style" }
                    }
                    '^(Scope|Applicability):' { 
                        try { $para.Style = "Section Heading" } catch { Write-Warning "Could not set Section Heading style" }
                    }
                    '^(Procedure|Process):' { 
                        try { $para.Style = "Subsection Heading" } catch { Write-Warning "Could not set Subsection Heading style" }
                    }
                }

                # Paragraph Formatting
                $paraFormat = $para.Format
                $paraFormat.LineSpacingRule = 1  # Single spacing
                $paraFormat.LeftIndent = 0
                $paraFormat.FirstLineIndent = 0
                $paraFormat.Alignment = 0  # Left aligned
                $paraFormat.SpaceBefore = 6
                $paraFormat.SpaceAfter = 6
            } catch {
                Write-Warning "Could not format paragraph $i : $_"
            }
        }

        # Safe Table Formatting
        try {
            foreach ($table in $doc.Tables) {
                try {
                    $table.Style = "Grid Table 4 - Accent 1"
                    $table.AutoFitBehavior(1)
                } catch {
                    Write-Warning "Could not format table: $_"
                }
            }
        } catch {
            Write-Warning "Table iteration failed: $_"
        }

        # Save Document
        try {
            $doc.SaveAs([ref]$OutputPath)
            Write-Host "CMMI Level 7 SOP formatted successfully: $OutputPath" -ForegroundColor Green
        } catch {
            throw "Could not save document: $_"
        }

    } catch {
        Write-Error "Formatting Error: $_"
    } finally {
        # Comprehensive Cleanup
        try {
            if ($doc) { 
                $doc.Close($false) 
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
            }
            if ($word) { 
                $word.Quit() 
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            }
        } catch {
            Write-Warning "Cleanup encountered an issue: $_"
        }

        # Force Garbage Collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Paths
$inputPath = "C:\Users\sm520\Desktop\SOP - ADFS Certificate Renewal_v0.1 - for merge.docx"
$outputPath = "C:\temp\CMMI_Level7_SOP_Certificate_Renewal.docx"

# Execute Formatting
Format-CMMI7SOP -InputPath $inputPath -OutputPath $outputPath
