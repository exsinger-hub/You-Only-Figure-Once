param(
    [Parameter(Mandatory = $true)]
    [string]$PayloadBase64
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Test-Property {
    param($Object, [string]$Name)
    return $null -ne $Object.PSObject.Properties[$Name]
}

function Get-Argument {
    param($Object, [string]$Name, $Default = $null)
    if (Test-Property $Object $Name) {
        return $Object.PSObject.Properties[$Name].Value
    }
    return $Default
}

$script:FocusPolicy = "preserve"

function Normalize-FocusPolicy {
    param($Value)
    $normalized = [string]$Value
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "preserve" }
    $normalized = $normalized.Trim().ToLowerInvariant()
    if ($normalized -notin @("preserve", "foreground")) { return "preserve" }
    return $normalized
}

function Initialize-FocusInterop {
    if (-not ("ScientificIllustrator.FocusWindow" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace ScientificIllustrator {
    public static class FocusWindow {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
"@
    }
}

function Get-ForegroundWindowHandle {
    Initialize-FocusInterop
    return [ScientificIllustrator.FocusWindow]::GetForegroundWindow()
}

function Restore-ForegroundWindow {
    param([IntPtr]$WindowHandle)
    if ($WindowHandle -eq [IntPtr]::Zero) { return }
    Initialize-FocusInterop
    if ([ScientificIllustrator.FocusWindow]::IsWindow($WindowHandle)) {
        $null = [ScientificIllustrator.FocusWindow]::SetForegroundWindow($WindowHandle)
    }
}

function Import-OfficeInteropMetadata {
    $result = [ordered]@{
        office_core = $false
        powerpoint = $false
        excel_chart_types = $false
        errors = @()
    }
    try {
        Add-Type -AssemblyName office -ErrorAction Stop
        $result.office_core = $true
    }
    catch { $result.errors += "office: $($_.Exception.Message)" }
    try {
        Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint -ErrorAction Stop
        $result.powerpoint = $true
    }
    catch { $result.errors += "powerpoint: $($_.Exception.Message)" }
    try {
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel -ErrorAction Stop
        $result.excel_chart_types = $true
    }
    catch { $result.errors += "excel chart types: $($_.Exception.Message)" }
    return $result
}

function Convert-EnumNameToPluginName {
    param([string]$Name, [string]$Prefix)
    $core = $Name
    if (-not [string]::IsNullOrWhiteSpace($Prefix) -and $core.StartsWith($Prefix, [StringComparison]::OrdinalIgnoreCase)) {
        $core = $core.Substring($Prefix.Length)
    }
    $core = [regex]::Replace($core, "([A-Z]+)([A-Z][a-z])", '$1_$2')
    $core = [regex]::Replace($core, "([a-z0-9])([A-Z])", '$1_$2')
    $core = [regex]::Replace($core, "[^A-Za-z0-9]+", "_").Trim("_")
    return $core.ToLowerInvariant()
}

function Get-AutoShapeCategory {
    param([string]$PluginName)
    if ($PluginName -like "flowchart_*") { return "flowchart" }
    if ($PluginName -like "action_button_*") { return "action_button" }
    if ($PluginName -match "callout") { return "callout" }
    if ($PluginName -match "arrow") { return "arrow" }
    if ($PluginName -match "star|explosion|sun|moon|smiley|heart|lightning") { return "star_and_banner" }
    if ($PluginName -match "brace|bracket") { return "brace_and_bracket" }
    if ($PluginName -match "arc|pie|chord|wave") { return "curve" }
    return "basic_shape"
}

function Get-EnumCatalog {
    param($EnumType, [string]$Prefix, [string]$CatalogKind)
    $items = @()
    foreach ($enumName in [Enum]::GetNames($EnumType)) {
        $pluginName = Convert-EnumNameToPluginName $enumName $Prefix
        $item = [ordered]@{
            plugin_name = $pluginName
            office_name = [string]$enumName
            value = [int][Enum]::Parse($EnumType, $enumName)
        }
        if ($CatalogKind -eq "auto_shape") { $item.category = Get-AutoShapeCategory $pluginName }
        $items += [pscustomobject]$item
    }
    return @($items | Sort-Object value, office_name)
}

function Invoke-Capabilities {
    param($Arguments)
    $installed = $false
    try { $installed = Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID" } catch {}
    $processes = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    $application = Get-PowerPointApplication $false
    $interop = Import-OfficeInteropMetadata

    $shapeCollectionMethods = @()
    $shapeMethods = @()
    $shapeRangeMethods = @()
    if ($interop.powerpoint) {
        $shapeCollectionMethods = @([Microsoft.Office.Interop.PowerPoint.Shapes].GetMethods() | Select-Object -ExpandProperty Name -Unique | Sort-Object)
        $shapeMethods = @([Microsoft.Office.Interop.PowerPoint.Shape].GetMethods() | Select-Object -ExpandProperty Name -Unique | Sort-Object)
        $shapeRangeMethods = @([Microsoft.Office.Interop.PowerPoint.ShapeRange].GetMethods() | Select-Object -ExpandProperty Name -Unique | Sort-Object)
    }

    $hasAddChart = ($shapeCollectionMethods -contains "AddChart2") -or ($shapeCollectionMethods -contains "AddChart")
    $families = @(
        [pscustomobject][ordered]@{ family = "text_box"; powerpoint_api = "Shapes.AddTextbox"; host_supported = $shapeCollectionMethods -contains "AddTextbox"; editable = $true; preferred_for = @("titles", "labels", "captions", "paragraphs") },
        [pscustomobject][ordered]@{ family = "auto_shape"; powerpoint_api = "Shapes.AddShape"; host_supported = $shapeCollectionMethods -contains "AddShape"; editable = $true; preferred_for = @("rectangles", "rounded boxes", "symbols", "block arrows", "flowchart nodes") },
        [pscustomobject][ordered]@{ family = "free_line_or_arrow"; powerpoint_api = "Shapes.AddLine"; host_supported = $shapeCollectionMethods -contains "AddLine"; editable = $true; preferred_for = @("arrows", "axes", "ticks", "separators", "annotations") },
        [pscustomobject][ordered]@{ family = "attached_connector"; powerpoint_api = "Shapes.AddConnector"; host_supported = $shapeCollectionMethods -contains "AddConnector"; editable = $true; preferred_for = @("semantic links that must stay attached when nodes move") },
        [pscustomobject][ordered]@{ family = "table"; powerpoint_api = "Shapes.AddTable"; host_supported = $shapeCollectionMethods -contains "AddTable"; editable = $true; preferred_for = @("tables", "matrix layouts", "grid annotations") },
        [pscustomobject][ordered]@{ family = "chart"; powerpoint_api = "Shapes.AddChart2/AddChart"; host_supported = $hasAddChart; editable = $true; preferred_for = @("bar", "line", "scatter", "pie", "area", "regular quantitative plots") },
        [pscustomobject][ordered]@{ family = "picture_or_svg"; powerpoint_api = "Shapes.AddPicture"; host_supported = $shapeCollectionMethods -contains "AddPicture"; editable = $true; preferred_for = @("tightly cropped microscopy", "photographic texture", "heatmaps", "irreducible raster evidence"); restriction = "Never use for a whole panel containing reconstructable text, shapes, arrows, tables, charts, or legends." },
        [pscustomobject][ordered]@{ family = "freeform"; powerpoint_api = "Shapes.BuildFreeform"; host_supported = $shapeCollectionMethods -contains "BuildFreeform"; editable = $true; preferred_for = @("custom irregular vector outlines") },
        [pscustomobject][ordered]@{ family = "curve_or_polyline"; powerpoint_api = "Shapes.AddCurve/AddPolyline"; host_supported = (($shapeCollectionMethods -contains "AddCurve") -or ($shapeCollectionMethods -contains "AddPolyline")); editable = $true; preferred_for = @("custom paths", "traces") },
        [pscustomobject][ordered]@{ family = "smartart"; powerpoint_api = "Shapes.AddSmartArt"; host_supported = $shapeCollectionMethods -contains "AddSmartArt"; editable = $true; preferred_for = @("built-in semantic process layouts") },
        [pscustomobject][ordered]@{ family = "duplicate"; powerpoint_api = "Shape.Duplicate"; host_supported = $shapeMethods -contains "Duplicate"; editable = $true; preferred_for = @("repeated native motifs") },
        [pscustomobject][ordered]@{ family = "group"; powerpoint_api = "ShapeRange.Group"; host_supported = $shapeRangeMethods -contains "Group"; editable = $true; preferred_for = @("panel-local object groups") },
        [pscustomobject][ordered]@{ family = "ungroup"; powerpoint_api = "Shape.Ungroup"; host_supported = $shapeMethods -contains "Ungroup"; editable = $true; preferred_for = @("editing members of an existing group") },
        [pscustomobject][ordered]@{ family = "z_order"; powerpoint_api = "Shape.ZOrder"; host_supported = $shapeMethods -contains "ZOrder"; editable = $true; preferred_for = @("layering", "backgrounds", "overlays") },
        [pscustomobject][ordered]@{ family = "align"; powerpoint_api = "ShapeRange.Align"; host_supported = $shapeRangeMethods -contains "Align"; editable = $true; preferred_for = @("shared edges", "shared centers", "regular rows and columns") },
        [pscustomobject][ordered]@{ family = "distribute"; powerpoint_api = "ShapeRange.Distribute"; host_supported = $shapeRangeMethods -contains "Distribute"; editable = $true; preferred_for = @("equal horizontal gaps", "equal vertical gaps", "repeated motifs") },
        [pscustomobject][ordered]@{ family = "figure_audit"; powerpoint_api = "Scientific Illustrator structure and renderer audit"; host_supported = $true; editable = $false; preferred_for = @("text fit", "connector clearance", "repeated layout", "atomic raster review") },
        [pscustomobject][ordered]@{ family = "media_or_ole"; powerpoint_api = "Shapes.AddMediaObject2/AddOLEObject"; host_supported = (($shapeCollectionMethods -contains "AddMediaObject2") -or ($shapeCollectionMethods -contains "AddOLEObject")); editable = $true; preferred_for = @("embedded media or external objects") }
    )

    $hostInfo = [ordered]@{
        platform = "win32"
        installed = [bool]$installed
        running_processes = [int]$processes.Count
        process_ids = @($processes | ForEach-Object { [int]$_.Id })
        connected_to_active_application = $null -ne $application
        active_application_process_id = if ($null -ne $application) { Get-PowerPointProcessId $application } else { 0 }
        active_presentation = $false
        application_version = ""
        application_build = ""
    }
    if ($null -ne $application) {
        try { $hostInfo.application_version = [string]$application.Version } catch {}
        try { $hostInfo.application_build = [string]$application.Build } catch {}
        try { $hostInfo.active_presentation = $null -ne $application.ActivePresentation } catch {}
    }

    $result = [ordered]@{
        detection = [ordered]@{
            read_only = $true
            launched_powerpoint = $false
            active_deck_modified = $false
            basis = @("installed Office interop type metadata", "running PowerPoint COM host metadata when available")
        }
        host = $hostInfo
        interop_metadata = $interop
        native_object_families = $families
        connector_types = @()
        arrowhead_styles = @()
        line_dash_styles = @()
        z_order_commands = @()
    }
    if ($interop.office_core) {
        $result.connector_types = Get-EnumCatalog ([Microsoft.Office.Core.MsoConnectorType]) "msoConnector" "connector"
        $result.arrowhead_styles = Get-EnumCatalog ([Microsoft.Office.Core.MsoArrowheadStyle]) "msoArrowhead" "arrowhead"
        $result.line_dash_styles = Get-EnumCatalog ([Microsoft.Office.Core.MsoLineDashStyle]) "msoLine" "line_dash"
        $result.z_order_commands = Get-EnumCatalog ([Microsoft.Office.Core.MsoZOrderCmd]) "mso" "z_order"
        if ([bool](Get-Argument $Arguments "include_auto_shapes" $true)) {
            $result.auto_shapes = Get-EnumCatalog ([Microsoft.Office.Core.MsoAutoShapeType]) "msoShape" "auto_shape"
        }
        if ([bool](Get-Argument $Arguments "include_shape_types" $true)) {
            $result.shape_types = Get-EnumCatalog ([Microsoft.Office.Core.MsoShapeType]) "mso" "shape_type"
        }
    }
    if ($interop.excel_chart_types -and [bool](Get-Argument $Arguments "include_chart_types" $true)) {
        $result.chart_types = Get-EnumCatalog ([Microsoft.Office.Interop.Excel.XlChartType]) "xl" "chart_type"
    }
    if ([bool](Get-Argument $Arguments "include_api_methods" $false)) {
        $result.api_methods = [ordered]@{
            shapes_collection = $shapeCollectionMethods
            shape = $shapeMethods
            shape_range = $shapeRangeMethods
        }
    }
    return $result
}

function Convert-HexToOfficeRgb {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Color must not be empty."
    }
    $hex = $Value.Trim().TrimStart("#")
    if ($hex -notmatch "^[0-9A-Fa-f]{6}$") {
        throw "Invalid color '$Value'; expected #RRGGBB."
    }
    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return [int]($red -bor ($green -shl 8) -bor ($blue -shl 16))
}

function Get-PowerPointApplication {
    param([bool]$Create)
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    }
    catch {
        if (-not $Create) {
            return $null
        }
        try {
            return New-Object -ComObject PowerPoint.Application
        }
        catch {
            throw "Unable to start desktop PowerPoint: $($_.Exception.Message)"
        }
    }
}

function Get-PowerPointProcessId {
    param($Application)
    if ($null -eq $Application) { return 0 }
    if ($null -eq ("ScientificIllustrator.NativeWindow" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace ScientificIllustrator {
    public static class NativeWindow {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
"@
    }
    $processIdValue = [uint32]0
    try { $null = [ScientificIllustrator.NativeWindow]::GetWindowThreadProcessId([IntPtr][int64]$Application.HWND, [ref]$processIdValue) } catch {}
    if ($processIdValue -gt 0) { return [int]$processIdValue }
    $powerPointProcesses = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    if ($powerPointProcesses.Count -eq 1) { return [int]$powerPointProcesses[0].Id }
    return 0
}

function Get-ActivePresentation {
    param($Application)
    if ($null -eq $Application) {
        throw "PowerPoint is not running. Call powerpoint_launch first."
    }
    try {
        $presentation = $Application.ActivePresentation
    }
    catch {
        $presentation = $null
    }
    if ($null -eq $presentation) {
        throw "PowerPoint has no active presentation. Call powerpoint_launch with a file or create_if_missing=true."
    }
    return $presentation
}

function Get-Slide {
    param($Presentation, [int]$Index)
    if ($Index -lt 1 -or $Index -gt $Presentation.Slides.Count) {
        throw "slide_index $Index is outside the valid range 1..$($Presentation.Slides.Count)."
    }
    return $Presentation.Slides.Item($Index)
}

function Show-Slide {
    param($Application, [int]$Index, [bool]$ForceForeground = $false)
    try {
        if ($null -ne $Application.ActiveWindow) {
            $Application.ActiveWindow.ViewType = 9
            $Application.ActiveWindow.View.GotoSlide($Index)
            if ($ForceForeground -or $script:FocusPolicy -eq "foreground") {
                $Application.ActiveWindow.Activate()
            }
        }
    }
    catch {
        # The edit remains valid when PowerPoint is in a transient modal/view state.
    }
}

function Find-Shape {
    param($Slide, $Arguments)
    $shapeName = Get-Argument $Arguments "shape_name"
    $shapeId = Get-Argument $Arguments "shape_id"
    $matches = @()
    foreach ($shape in @($Slide.Shapes)) {
        if ($null -ne $shapeName -and $shape.Name -ieq [string]$shapeName) {
            $matches += $shape
        }
        elseif ($null -ne $shapeId -and [int]$shape.Id -eq [int]$shapeId) {
            $matches += $shape
        }
    }
    if ($matches.Count -gt 1) { throw "Shape target is ambiguous because semantic name '$shapeName' occurs $($matches.Count) times on slide $($Slide.SlideIndex)." }
    if ($matches.Count -eq 1) { return $matches[0] }
    if ($null -ne $shapeName) {
        throw "Shape '$shapeName' was not found on slide $($Slide.SlideIndex)."
    }
    if ($null -ne $shapeId) {
        throw "Shape id $shapeId was not found on slide $($Slide.SlideIndex)."
    }
    throw "Provide shape_name or shape_id."
}

function Assert-ShapeNameAvailable {
    param($Slide, [string]$Name, $ExceptShape = $null)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }
    foreach ($shape in @($Slide.Shapes)) {
        if ($shape.Name -ieq $Name -and ($null -eq $ExceptShape -or $shape.Id -ne $ExceptShape.Id)) {
            throw "Shape name '$Name' already exists on slide $($Slide.SlideIndex)."
        }
    }
}

function Get-ShapeText {
    param($Shape)
    try {
        if ($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1) {
            return [string]$Shape.TextFrame.TextRange.Text
        }
    }
    catch {}
    return ""
}

function Get-ShapeTag {
    param($Shape, [string]$Name)
    try { return [string]$Shape.Tags.Item($Name) } catch { return "" }
}

function New-ShapeSummary {
    param($Shape)
    $altText = ""
    try { $altText = [string]$Shape.AlternativeText } catch {}
    $shapeTypeName = ""
    $autoShapeTypeName = ""
    $zOrderPosition = 0
    $isTable = $false
    $isChart = $false
    $groupItemCount = 0
    try {
        Add-Type -AssemblyName office -ErrorAction Stop
        $shapeTypeName = [string][Enum]::GetName([Microsoft.Office.Core.MsoShapeType], [int]$Shape.Type)
        if ([int]$Shape.Type -eq 1) {
            $autoShapeTypeName = [string][Enum]::GetName([Microsoft.Office.Core.MsoAutoShapeType], [int]$Shape.AutoShapeType)
        }
    }
    catch {}
    try { $zOrderPosition = [int]$Shape.ZOrderPosition } catch {}
    try { $isTable = $Shape.HasTable -eq -1 } catch {}
    try { $isChart = $Shape.HasChart -eq -1 } catch {}
    try { $groupItemCount = [int]$Shape.GroupItems.Count } catch {}
    $summary = [ordered]@{
        id = [int]$Shape.Id
        name = [string]$Shape.Name
        type = [int]$Shape.Type
        type_name = $shapeTypeName
        auto_shape_type = $autoShapeTypeName
        left = [math]::Round([double]$Shape.Left, 2)
        top = [math]::Round([double]$Shape.Top, 2)
        width = [math]::Round([double]$Shape.Width, 2)
        height = [math]::Round([double]$Shape.Height, 2)
        rotation = [math]::Round([double]$Shape.Rotation, 2)
        z_order_position = $zOrderPosition
        text = Get-ShapeText $Shape
        alt_text = $altText
        is_table = $isTable
        is_chart = $isChart
        group_item_count = $groupItemCount
    }
    $rasterReason = Get-ShapeTag $Shape "ScientificIllustratorRasterReason"
    $sourceTight = Get-ShapeTag $Shape "ScientificIllustratorSourceTightlyCropped"
    $atomicRaster = Get-ShapeTag $Shape "ScientificIllustratorAtomicRasterUnit"
    $containsReconstructable = Get-ShapeTag $Shape "ScientificIllustratorContainsReconstructableContent"
    $decompositionNote = Get-ShapeTag $Shape "ScientificIllustratorDecompositionNote"
    if (-not [string]::IsNullOrWhiteSpace($rasterReason) -or [int]$Shape.Type -eq 13) {
        $summary.raster_reason = $rasterReason
        $summary.source_is_tightly_cropped = $sourceTight
        $summary.atomic_raster_unit = $atomicRaster
        $summary.contains_reconstructable_content = $containsReconstructable
        $summary.decomposition_note = $decompositionNote
    }
    try {
        if ($Shape.HasTextFrame -eq -1 -and $Shape.TextFrame.HasText -eq -1) {
            $boundWidth = [double]$Shape.TextFrame2.TextRange.BoundWidth
            $boundHeight = [double]$Shape.TextFrame2.TextRange.BoundHeight
            $marginLeft = [double]$Shape.TextFrame.MarginLeft
            $marginRight = [double]$Shape.TextFrame.MarginRight
            $marginTop = [double]$Shape.TextFrame.MarginTop
            $marginBottom = [double]$Shape.TextFrame.MarginBottom
            $innerWidth = [math]::Max(0, [double]$Shape.Width - $marginLeft - $marginRight)
            $innerHeight = [math]::Max(0, [double]$Shape.Height - $marginTop - $marginBottom)
            $summary.text_frame = [ordered]@{
                margin_left = [math]::Round($marginLeft, 2)
                margin_right = [math]::Round($marginRight, 2)
                margin_top = [math]::Round($marginTop, 2)
                margin_bottom = [math]::Round($marginBottom, 2)
                word_wrap = [int]$Shape.TextFrame.WordWrap
                auto_size = [int]$Shape.TextFrame2.AutoSize
                bound_width = [math]::Round($boundWidth, 2)
                bound_height = [math]::Round($boundHeight, 2)
                inner_width = [math]::Round($innerWidth, 2)
                inner_height = [math]::Round($innerHeight, 2)
            }
        }
    }
    catch {}
    try {
        if ([int]$Shape.Type -eq 9) {
            $beginX = [double]$Shape.Left
            $beginY = [double]$Shape.Top
            $endX = [double]$Shape.Left + [double]$Shape.Width
            $endY = [double]$Shape.Top + [double]$Shape.Height
            $tagBeginX = Get-ShapeTag $Shape "ScientificIllustratorBeginX"
            $tagBeginY = Get-ShapeTag $Shape "ScientificIllustratorBeginY"
            $tagEndX = Get-ShapeTag $Shape "ScientificIllustratorEndX"
            $tagEndY = Get-ShapeTag $Shape "ScientificIllustratorEndY"
            if (-not [string]::IsNullOrWhiteSpace($tagBeginX)) { $beginX = [double]$tagBeginX }
            if (-not [string]::IsNullOrWhiteSpace($tagBeginY)) { $beginY = [double]$tagBeginY }
            if (-not [string]::IsNullOrWhiteSpace($tagEndX)) { $endX = [double]$tagEndX }
            if (-not [string]::IsNullOrWhiteSpace($tagEndY)) { $endY = [double]$tagEndY }
            $connectorType = 0
            try { if ($Shape.Connector -eq -1) { $connectorType = [int]$Shape.ConnectorFormat.Type } } catch {}
            $summary.line = [ordered]@{
                begin_x = [math]::Round($beginX, 2)
                begin_y = [math]::Round($beginY, 2)
                end_x = [math]::Round($endX, 2)
                end_y = [math]::Round($endY, 2)
                begin_arrow = [int]$Shape.Line.BeginArrowheadStyle
                end_arrow = [int]$Shape.Line.EndArrowheadStyle
                connector = [int]$Shape.Connector
                connector_type = $connectorType
                source_name = Get-ShapeTag $Shape "ScientificIllustratorSourceName"
                target_name = Get-ShapeTag $Shape "ScientificIllustratorTargetName"
                start_clearance = Get-ShapeTag $Shape "ScientificIllustratorStartClearance"
                end_clearance = Get-ShapeTag $Shape "ScientificIllustratorEndClearance"
            }
        }
    }
    catch {}
    if ($isTable) {
        try {
            $summary.table_rows = [int]$Shape.Table.Rows.Count
            $summary.table_columns = [int]$Shape.Table.Columns.Count
        }
        catch {}
    }
    return $summary
}

function Set-ShapeGeometry {
    param($Shape, $Arguments)
    if (Test-Property $Arguments "left") { $Shape.Left = [single](Get-Argument $Arguments "left") }
    if (Test-Property $Arguments "top") { $Shape.Top = [single](Get-Argument $Arguments "top") }
    if (Test-Property $Arguments "width") { $Shape.Width = [single](Get-Argument $Arguments "width") }
    if (Test-Property $Arguments "height") { $Shape.Height = [single](Get-Argument $Arguments "height") }
    if (Test-Property $Arguments "rotation") { $Shape.Rotation = [single](Get-Argument $Arguments "rotation") }
}

function Set-ShapeAppearance {
    param($Shape, $Arguments)
    if (Test-Property $Arguments "fill_color") {
        $Shape.Fill.Visible = -1
        $Shape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb (Get-Argument $Arguments "fill_color")
        $Shape.Fill.Solid()
    }
    if (Test-Property $Arguments "fill_transparency") {
        $Shape.Fill.Transparency = [single]([double](Get-Argument $Arguments "fill_transparency") / 100.0)
    }
    if (Test-Property $Arguments "line_width") {
        $lineWidth = [double](Get-Argument $Arguments "line_width")
        if ($lineWidth -eq 0) {
            $Shape.Line.Visible = 0
        }
        else {
            $Shape.Line.Visible = -1
            $Shape.Line.Weight = [single]$lineWidth
        }
    }
    if (Test-Property $Arguments "line_color") {
        $Shape.Line.Visible = -1
        $Shape.Line.ForeColor.RGB = Convert-HexToOfficeRgb (Get-Argument $Arguments "line_color")
    }
    if (Test-Property $Arguments "line_transparency") {
        $Shape.Line.Visible = -1
        $Shape.Line.Transparency = [single]([double](Get-Argument $Arguments "line_transparency") / 100.0)
    }
    if (Test-Property $Arguments "line_dash") {
        $dashMap = @{
            solid = 1; square_dot = 2; round_dot = 3; dash = 4; dash_dot = 5
            long_dash = 6; long_dash_dot = 7; long_dash_dot_dot = 8
        }
        $Shape.Line.Visible = -1
        $Shape.Line.DashStyle = $dashMap[[string](Get-Argument $Arguments "line_dash")]
    }
    if (Test-Property $Arguments "start_arrow") {
        $Shape.Line.BeginArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "start_arrow"))
    }
    if (Test-Property $Arguments "end_arrow") {
        $Shape.Line.EndArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "end_arrow"))
    }
}

function Set-ShapeText {
    param($Shape, $Arguments)
    $hasTextSettings = (Test-Property $Arguments "text") -or
        (Test-Property $Arguments "font_name") -or
        (Test-Property $Arguments "font_size") -or
        (Test-Property $Arguments "font_color") -or
        (Test-Property $Arguments "bold") -or
        (Test-Property $Arguments "italic") -or
        (Test-Property $Arguments "alignment") -or
        (Test-Property $Arguments "vertical_alignment") -or
        (Test-Property $Arguments "margin_left") -or
        (Test-Property $Arguments "margin_right") -or
        (Test-Property $Arguments "margin_top") -or
        (Test-Property $Arguments "margin_bottom") -or
        (Test-Property $Arguments "word_wrap") -or
        (Test-Property $Arguments "text_autofit")
    if (-not $hasTextSettings) {
        return
    }
    if ($Shape.HasTextFrame -ne -1) {
        throw "Shape '$($Shape.Name)' does not support text."
    }
    if (Test-Property $Arguments "text") {
        $Shape.TextFrame.TextRange.Text = [string](Get-Argument $Arguments "text")
    }
    $range = $Shape.TextFrame.TextRange
    if (Test-Property $Arguments "font_name") { $range.Font.Name = [string](Get-Argument $Arguments "font_name") }
    if (Test-Property $Arguments "font_size") { $range.Font.Size = [single](Get-Argument $Arguments "font_size") }
    if (Test-Property $Arguments "font_color") { $range.Font.Color.RGB = Convert-HexToOfficeRgb (Get-Argument $Arguments "font_color") }
    if (Test-Property $Arguments "bold") { $range.Font.Bold = if ([bool](Get-Argument $Arguments "bold")) { -1 } else { 0 } }
    if (Test-Property $Arguments "italic") { $range.Font.Italic = if ([bool](Get-Argument $Arguments "italic")) { -1 } else { 0 } }
    if (Test-Property $Arguments "alignment") {
        $alignmentMap = @{ left = 1; center = 2; right = 3; justify = 4 }
        $alignment = [string](Get-Argument $Arguments "alignment")
        $range.ParagraphFormat.Alignment = $alignmentMap[$alignment]
    }
    if (Test-Property $Arguments "vertical_alignment") {
        $verticalMap = @{ top = 1; middle = 3; bottom = 4 }
        $vertical = [string](Get-Argument $Arguments "vertical_alignment")
        $Shape.TextFrame.VerticalAnchor = $verticalMap[$vertical]
    }
    if (Test-Property $Arguments "word_wrap") {
        $Shape.TextFrame.WordWrap = if ([bool](Get-Argument $Arguments "word_wrap")) { -1 } else { 0 }
    }
    if (Test-Property $Arguments "text_autofit") {
        $autofit = [string](Get-Argument $Arguments "text_autofit" "none")
        if ($autofit -eq "grow_shape") {
            $Shape.TextFrame.AutoSize = 1
            try { $Shape.TextFrame2.AutoSize = 1 } catch {}
        }
        elseif ($autofit -eq "shrink_text") {
            $Shape.TextFrame.AutoSize = 0
            try { $Shape.TextFrame2.AutoSize = 2 } catch {}
        }
        else {
            $Shape.TextFrame.AutoSize = 0
            try { $Shape.TextFrame2.AutoSize = 0 } catch {}
        }
    }
    foreach ($margin in @("margin_left", "margin_right", "margin_top", "margin_bottom")) {
        if (Test-Property $Arguments $margin) {
            $propertyName = switch ($margin) {
                "margin_left" { "MarginLeft" }
                "margin_right" { "MarginRight" }
                "margin_top" { "MarginTop" }
                "margin_bottom" { "MarginBottom" }
            }
            $Shape.TextFrame.$propertyName = [single](Get-Argument $Arguments $margin)
        }
    }
}

function Set-ShapeFromArguments {
    param($Shape, $Arguments)
    Set-ShapeGeometry $Shape $Arguments
    Set-ShapeAppearance $Shape $Arguments
    Set-ShapeText $Shape $Arguments
}

function Get-PresentationSummary {
    param($Application, $Presentation)
    $fullName = ""
    $saved = $false
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$Presentation.Path)) {
            $fullName = [string]$Presentation.FullName
        }
    }
    catch {}
    try { $saved = $Presentation.Saved -eq -1 } catch {}
    return [ordered]@{
        application_version = [string]$Application.Version
        presentation_name = [string]$Presentation.Name
        presentation_path = $fullName
        has_file_path = -not [string]::IsNullOrWhiteSpace($fullName)
        slide_count = [int]$Presentation.Slides.Count
        slide_width = [math]::Round([double]$Presentation.PageSetup.SlideWidth, 2)
        slide_height = [math]::Round([double]$Presentation.PageSetup.SlideHeight, 2)
        saved = $saved
    }
}

function Invoke-NewPresentation {
    param($Arguments)
    $application = Get-PowerPointApplication $true
    $application.Visible = -1
    $presentation = $application.Presentations.Add($true)
    if ($script:FocusPolicy -eq "foreground") {
        try { $presentation.Windows.Item(1).Activate() } catch {}
    }
    if ([bool](Get-Argument $Arguments "maximize" $true)) {
        try { $application.ActiveWindow.WindowState = 3 } catch {}
    }
    $summary = Get-PresentationSummary $application $presentation
    $summary.created = $true
    $summary.connected = $true
    return $summary
}

function Invoke-Status {
    $installed = $false
    try { $installed = Test-Path "Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID" } catch {}
    $processes = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
    $application = Get-PowerPointApplication $false
    $activePresentation = $null
    if ($null -ne $application) {
        try { $activePresentation = $application.ActivePresentation } catch {}
    }
    $result = [ordered]@{
        platform = "win32"
        installed = [bool]$installed
        running_processes = [int]$processes.Count
        process_ids = @($processes | ForEach-Object { [int]$_.Id })
        connected_to_active_application = $null -ne $application
        active_application_process_id = if ($null -ne $application) { Get-PowerPointProcessId $application } else { 0 }
        active_presentation = $null -ne $activePresentation
        control_scope = "PowerPoint native COM object model only"
        focus_policy = $script:FocusPolicy
    }
    if ($null -ne $activePresentation) {
        $result.presentation = Get-PresentationSummary $application $activePresentation
    }
    return $result
}

function Invoke-Launch {
    param($Arguments)
    $application = Get-PowerPointApplication $true
    $visible = [bool](Get-Argument $Arguments "visible" $true)
    $application.Visible = if ($visible) { -1 } else { 0 }
    $filePath = Get-Argument $Arguments "file_path"
    $presentation = $null

    if (-not [string]::IsNullOrWhiteSpace([string]$filePath)) {
        $resolved = [IO.Path]::GetFullPath([string]$filePath)
        if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
            throw "PowerPoint file does not exist: $resolved"
        }
        foreach ($openPresentation in @($application.Presentations)) {
            try {
                if ([IO.Path]::GetFullPath([string]$openPresentation.FullName) -ieq $resolved) {
                    $presentation = $openPresentation
                    break
                }
            }
            catch {}
        }
        if ($null -eq $presentation) {
            $readOnly = [bool](Get-Argument $Arguments "read_only" $false)
            $presentation = $application.Presentations.Open($resolved, $readOnly, $false, $visible)
        }
    }
    else {
        try { $presentation = $application.ActivePresentation } catch {}
        if ($null -eq $presentation -and [bool](Get-Argument $Arguments "create_if_missing" $true)) {
            $presentation = $application.Presentations.Add($visible)
        }
    }

    if ($null -eq $presentation) {
        throw "No active presentation is available and create_if_missing=false."
    }
    if ($script:FocusPolicy -eq "foreground") {
        try { $presentation.Windows.Item(1).Activate() } catch {}
    }
    if ([bool](Get-Argument $Arguments "maximize" $true)) {
        try { $application.ActiveWindow.WindowState = 3 } catch {}
    }
    $summary = Get-PresentationSummary $application $presentation
    $summary.connected = $true
    $summary.visible = $visible
    return $summary
}

function Invoke-Inspect {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $maxSlides = [int](Get-Argument $Arguments "max_slides" 100)
    $maxShapes = [int](Get-Argument $Arguments "max_shapes_per_slide" 200)
    $includeText = [bool](Get-Argument $Arguments "include_text" $true)
    $slideItems = @()
    $slideLimit = [math]::Min($presentation.Slides.Count, $maxSlides)
    for ($slideIndex = 1; $slideIndex -le $slideLimit; $slideIndex += 1) {
        $slide = $presentation.Slides.Item($slideIndex)
        $shapeItems = @()
        $shapeLimit = [math]::Min($slide.Shapes.Count, $maxShapes)
        for ($shapeIndex = 1; $shapeIndex -le $shapeLimit; $shapeIndex += 1) {
            $shape = $slide.Shapes.Item($shapeIndex)
            $shapeSummary = New-ShapeSummary $shape
            if (-not $includeText) { $shapeSummary.Remove("text") }
            $shapeItems += [pscustomobject]$shapeSummary
        }
        $slideName = ""
        try { $slideName = [string]$slide.Name } catch {}
        $slideItems += [pscustomobject][ordered]@{
            index = [int]$slideIndex
            id = [int]$slide.SlideID
            name = $slideName
            shape_count = [int]$slide.Shapes.Count
            shapes = $shapeItems
            shapes_truncated = $slide.Shapes.Count -gt $maxShapes
        }
    }
    $summary = Get-PresentationSummary $application $presentation
    $summary.slides = $slideItems
    $summary.slides_truncated = $presentation.Slides.Count -gt $maxSlides
    return $summary
}

function Invoke-AuditFigure {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $alignmentTolerance = [double](Get-Argument $Arguments "alignment_tolerance" 0.75)
    $endpointClearance = [double](Get-Argument $Arguments "endpoint_clearance" 1.5)
    $textTolerance = [double](Get-Argument $Arguments "text_overflow_tolerance" 1.5)
    $largeRasterRatio = [double](Get-Argument $Arguments "large_raster_area_ratio" 0.08)
    $maxFindings = [int](Get-Argument $Arguments "max_findings" 300)
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    $slideArea = [math]::Max(1, $slideWidth * $slideHeight)
    $findings = New-Object System.Collections.ArrayList
    $addFinding = {
        param([string]$Category, [string]$Severity, [object[]]$Objects, [string]$Evidence, [string]$Correction, [string]$Acceptance)
        if ($findings.Count -ge $maxFindings) { return }
        $null = $findings.Add([pscustomobject][ordered]@{
            category = $Category
            severity = $Severity
            objects = @($Objects)
            evidence = $Evidence
            correction = $Correction
            acceptance = $Acceptance
        })
    }
    $segmentIntersectsRect = {
        param([double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [double]$Inset)
        $margin = [math]::Max(0, $Inset)
        if ($Width -le (2 * $margin) -or $Height -le (2 * $margin)) { $margin = 0 }
        $rectLeft = $Left + $margin
        $rectRight = $Left + $Width - $margin
        $rectTop = $Top + $margin
        $rectBottom = $Top + $Height - $margin
        $insideFirst = $X1 -gt $rectLeft -and $X1 -lt $rectRight -and $Y1 -gt $rectTop -and $Y1 -lt $rectBottom
        $insideSecond = $X2 -gt $rectLeft -and $X2 -lt $rectRight -and $Y2 -gt $rectTop -and $Y2 -lt $rectBottom
        if ($insideFirst -or $insideSecond) { return $true }
        $deltaX = $X2 - $X1
        $deltaY = $Y2 - $Y1
        $p = @((-$deltaX), $deltaX, (-$deltaY), $deltaY)
        $q = @(($X1 - $rectLeft), ($rectRight - $X1), ($Y1 - $rectTop), ($rectBottom - $Y1))
        $minimum = 0.0
        $maximum = 1.0
        for ($index = 0; $index -lt 4; $index += 1) {
            if ([math]::Abs([double]$p[$index]) -lt 0.000000001) {
                if ([double]$q[$index] -lt 0) { return $false }
                continue
            }
            $ratio = [double]$q[$index] / [double]$p[$index]
            if ([double]$p[$index] -lt 0) { $minimum = [math]::Max($minimum, $ratio) }
            else { $maximum = [math]::Min($maximum, $ratio) }
            if ($minimum -gt $maximum) { return $false }
        }
        return ($maximum - $minimum) -gt 0.000001
    }
    $strictSegmentCrossing = {
        param([double]$AX1, [double]$AY1, [double]$AX2, [double]$AY2, [double]$BX1, [double]$BY1, [double]$BX2, [double]$BY2, [double]$Clearance)
        $aX = $AX2 - $AX1
        $aY = $AY2 - $AY1
        $bX = $BX2 - $BX1
        $bY = $BY2 - $BY1
        $denominator = $aX * $bY - $aY * $bX
        if ([math]::Abs($denominator) -lt 0.000000001) { return $null }
        $cX = $BX1 - $AX1
        $cY = $BY1 - $AY1
        $aRatio = ($cX * $bY - $cY * $bX) / $denominator
        $bRatio = ($cX * $aY - $cY * $aX) / $denominator
        if ($aRatio -le 0.000001 -or $aRatio -ge 0.999999 -or $bRatio -le 0.000001 -or $bRatio -ge 0.999999) { return $null }
        $aLength = [math]::Sqrt($aX * $aX + $aY * $aY)
        $bLength = [math]::Sqrt($bX * $bX + $bY * $bY)
        $distance = [math]::Min([math]::Min($aRatio * $aLength, (1 - $aRatio) * $aLength), [math]::Min($bRatio * $bLength, (1 - $bRatio) * $bLength))
        if ($distance -le $Clearance) { return $null }
        return [pscustomobject]@{ x = $AX1 + $aRatio * $aX; y = $AY1 + $aRatio * $aY }
    }

    $summaries = @()
    $shapes = @()
    $nameCounts = @{}
    for ($index = 1; $index -le $slide.Shapes.Count; $index += 1) {
        $shape = $slide.Shapes.Item($index)
        $shapes += $shape
        $summary = New-ShapeSummary $shape
        $summaries += [pscustomobject]$summary
        $key = ([string]$shape.Name).ToLowerInvariant()
        if (-not $nameCounts.ContainsKey($key)) { $nameCounts[$key] = 0 }
        $nameCounts[$key] += 1
        if ([double]$shape.Left -lt -$alignmentTolerance -or [double]$shape.Top -lt -$alignmentTolerance -or ([double]$shape.Left + [double]$shape.Width) -gt ($slideWidth + $alignmentTolerance) -or ([double]$shape.Top + [double]$shape.Height) -gt ($slideHeight + $alignmentTolerance)) {
            & $addFinding "outside-slide" "hard" @([string]$shape.Name) "Bounds [$([math]::Round($shape.Left,2)),$([math]::Round($shape.Top,2)),$([math]::Round($shape.Width,2)),$([math]::Round($shape.Height,2))] exceed $slideWidth x $slideHeight." "Move or resize the named object inside the slide." "All object bounds remain inside the slide within $alignmentTolerance pt."
        }
        try {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $boundWidth = [double]$shape.TextFrame2.TextRange.BoundWidth
                $boundHeight = [double]$shape.TextFrame2.TextRange.BoundHeight
                $innerWidth = [math]::Max(0, [double]$shape.Width - [double]$shape.TextFrame.MarginLeft - [double]$shape.TextFrame.MarginRight)
                $innerHeight = [math]::Max(0, [double]$shape.Height - [double]$shape.TextFrame.MarginTop - [double]$shape.TextFrame.MarginBottom)
                if ($boundWidth -gt ($innerWidth + $textTolerance) -or $boundHeight -gt ($innerHeight + $textTolerance)) {
                    & $addFinding "text-overflow" "hard" @([string]$shape.Name) "Text bounds $([math]::Round($boundWidth,2)) x $([math]::Round($boundHeight,2)) exceed inner frame $([math]::Round($innerWidth,2)) x $([math]::Round($innerHeight,2))." "Adjust the exact text box geometry, margins, line breaks, or text_autofit without rasterizing the label." "Text bounds fit the inner frame with at least $textTolerance pt tolerance and the renderer shows no clipping or unintended wrap."
                }
                if ([int]$shape.TextFrame2.AutoSize -eq 1) {
                    & $addFinding "layout-unstable-autofit" "warning" @([string]$shape.Name) "text_autofit=grow_shape can change the reference geometry after text is applied." "Use text_autofit=none for fixed reference reconstruction or shrink_text only when explicitly accepted." "The object retains its planned bounds after text updates."
                }
            }
        }
        catch {}
    }

    foreach ($entry in $nameCounts.GetEnumerator()) {
        if ([int]$entry.Value -gt 1) { & $addFinding "duplicate-name" "hard" @([string]$entry.Key) "The semantic name occurs $($entry.Value) times." "Rename every object uniquely before further correction calls." "Every object name on the slide is unique." }
    }

    $pictureAudits = @()
    foreach ($shape in $shapes) {
        if ([int]$shape.Type -notin @(11, 13)) { continue }
        $reason = Get-ShapeTag $shape "ScientificIllustratorRasterReason"
        $tight = Get-ShapeTag $shape "ScientificIllustratorSourceTightlyCropped"
        $atomic = Get-ShapeTag $shape "ScientificIllustratorAtomicRasterUnit"
        $containsNative = Get-ShapeTag $shape "ScientificIllustratorContainsReconstructableContent"
        $decomposition = Get-ShapeTag $shape "ScientificIllustratorDecompositionNote"
        $areaRatio = ([double]$shape.Width * [double]$shape.Height) / $slideArea
        $pictureAudit = [ordered]@{
            name = [string]$shape.Name
            area_ratio = [math]::Round($areaRatio, 4)
            raster_reason = $reason
            source_is_tightly_cropped = $tight
            atomic_raster_unit = $atomic
            contains_reconstructable_content = $containsNative
            decomposition_note = $decomposition
        }
        $pictureAudits += [pscustomobject]$pictureAudit
        if ([string]::IsNullOrWhiteSpace($reason)) { & $addFinding "raster-missing-reason" "hard" @([string]$shape.Name) "No serialized irreducibility reason is present." "Replace or retag the picture with a precise raster_reason." "The picture has a specific irreducibility reason." }
        if ($tight -ine "True") { & $addFinding "raster-not-tight" "hard" @([string]$shape.Name) "source_is_tightly_cropped is '$tight'." "Crop away every reconstructable border, label, arrow, legend, axis, or neighboring image." "The picture contains only its atomic visual field." }
        if ($atomic -ine "True") { & $addFinding "raster-not-atomic" "hard" @([string]$shape.Name) "atomic_raster_unit is '$atomic'." "Split the picture into one image object per microscopy field, mask, heatmap, photograph, or other irreducible datum." "Each retained picture is one indivisible raster unit." }
        if ($containsNative -ine "False") { & $addFinding "raster-contains-reconstructable-content" "hard" @([string]$shape.Name) "contains_reconstructable_content is '$containsNative'." "Rebuild all text, borders, arrows, legends, axes, tables, and regular plots as native objects." "No retained picture contains a reconstructable drawing primitive." }
        if ([string]::IsNullOrWhiteSpace($decomposition) -or $decomposition.Trim().Length -lt 8) { & $addFinding "raster-missing-decomposition-note" "hard" @([string]$shape.Name) "No useful decomposition note is serialized." "State what was separated and rebuilt natively, or why no finer semantic split is possible." "A reviewer can verify the atomic decomposition decision from the note." }
        if ($areaRatio -gt $largeRasterRatio) { & $addFinding "large-raster-surface" "warning" @([string]$shape.Name) "Picture occupies $([math]::Round($areaRatio * 100,1))% of the slide, above the $([math]::Round($largeRasterRatio * 100,1))% review threshold." "Visually inspect the source at full resolution and split any independent subimages or reconstructable overlay." "The reviewer confirms the large picture is still one atomic raster field." }
        $compositeText = "$($shape.Name) $reason $decomposition"
        if ($compositeText -match '(?i)(grid|montage|panel|comparison|stack|matrix|multi[- ]?image|multiple images|rows? of|columns? of)') {
            & $addFinding "possible-composite-raster" "hard" @([string]$shape.Name) "Name or reason suggests a composite raster: '$compositeText'." "Split each independent image, mask, heatmap, or error map into its own picture; recreate headings, grid, borders, and legend natively." "No picture name or reason describes a grid, montage, stack, panel, comparison, matrix, or multiple-image region."
        }
    }

    $candidateTargets = @($summaries | Where-Object {
        [int]$_.type -ne 9 -and [int]$_.type -ne 6 -and [string]$_.name -notmatch '(?i)(^|[-_])(panel|background|bg|container|region|frame)([-_]|$)' -and (([double]$_.width * [double]$_.height) / $slideArea) -lt 0.18 -and [double]$_.width -gt 2 -and [double]$_.height -gt 2
    } | Sort-Object { [double]$_.width * [double]$_.height })
    $straightArrowSegments = @()
    for ($shapeIndex = 0; $shapeIndex -lt $shapes.Count; $shapeIndex += 1) {
        $shape = $shapes[$shapeIndex]
        $summary = $summaries[$shapeIndex]
        if ([int]$summary.type -ne 9) { continue }
        if ($null -eq $summary.line) { continue }
        $sourceName = [string]$summary.line.source_name
        $targetName = [string]$summary.line.target_name
        $hasArrow = [int]$summary.line.begin_arrow -ne 1 -or [int]$summary.line.end_arrow -ne 1
        $isStraightRoute = $true
        try {
            if ($shape.Connector -eq -1 -and [int]$shape.ConnectorFormat.Type -ne 1) { $isStraightRoute = $false }
        }
        catch {}
        if ($hasArrow -and $isStraightRoute) {
            $segment = [pscustomobject][ordered]@{
                name = [string]$shape.Name
                source_name = $sourceName
                target_name = $targetName
                begin_x = [double]$summary.line.begin_x
                begin_y = [double]$summary.line.begin_y
                end_x = [double]$summary.line.end_x
                end_y = [double]$summary.line.end_y
            }
            $straightArrowSegments += $segment
            foreach ($target in $candidateTargets) {
                if ([int]$target.id -eq [int]$summary.id -or [string]$target.name -ieq $sourceName -or [string]$target.name -ieq $targetName) { continue }
                $intersects = & $segmentIntersectsRect $segment.begin_x $segment.begin_y $segment.end_x $segment.end_y ([double]$target.left) ([double]$target.top) ([double]$target.width) ([double]$target.height) $endpointClearance
                if ($intersects) {
                    & $addFinding "connector-path-through-object" "hard" @([string]$summary.name, [string]$target.name) "The straight rendered route of '$($summary.name)' passes through the interior of unrelated object '$($target.name)'." "Move the object or reroute the connector using explicit connection sites or an elbow connector in a reserved lane." "No connector segment intersects an unrelated shape or label interior."
                }
            }
        }
        $endpoints = @(
            [pscustomobject]@{ label = "start"; x = [double]$summary.line.begin_x; y = [double]$summary.line.begin_y; arrow = [int]$summary.line.begin_arrow; connected_name = $sourceName },
            [pscustomobject]@{ label = "end"; x = [double]$summary.line.end_x; y = [double]$summary.line.end_y; arrow = [int]$summary.line.end_arrow; connected_name = $targetName }
        )
        foreach ($endpoint in $endpoints) {
            if ([int]$endpoint.arrow -eq 1) { continue }
            foreach ($target in $candidateTargets) {
                if ([int]$target.id -eq [int]$summary.id -or ([string]$target.name -ieq [string]$endpoint.connected_name)) { continue }
                $inside = [double]$endpoint.x -gt ([double]$target.left + $endpointClearance) -and [double]$endpoint.x -lt ([double]$target.left + [double]$target.width - $endpointClearance) -and [double]$endpoint.y -gt ([double]$target.top + $endpointClearance) -and [double]$endpoint.y -lt ([double]$target.top + [double]$target.height - $endpointClearance)
                if ($inside) {
                    & $addFinding "arrowhead-intrusion" "hard" @([string]$summary.name, [string]$target.name) "$($endpoint.label) arrow endpoint ($([math]::Round($endpoint.x,2)),$([math]::Round($endpoint.y,2))) lies inside '$($target.name)' beyond $endpointClearance pt." "Attach a connector to an explicit connection site or trim the free line with start_clearance/end_clearance; then correct z-order." "The arrow tip touches the intended boundary without entering any unrelated object, and the shaft remains outside the fill."
                    break
                }
            }
        }
    }

    for ($firstIndex = 0; $firstIndex -lt $straightArrowSegments.Count; $firstIndex += 1) {
        $first = $straightArrowSegments[$firstIndex]
        for ($secondIndex = $firstIndex + 1; $secondIndex -lt $straightArrowSegments.Count; $secondIndex += 1) {
            $second = $straightArrowSegments[$secondIndex]
            $firstEndpoints = @($first.source_name, $first.target_name) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
            $secondEndpoints = @($second.source_name, $second.target_name) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
            if (@($firstEndpoints | Where-Object { $secondEndpoints -contains $_ }).Count -gt 0) { continue }
            $crossing = & $strictSegmentCrossing $first.begin_x $first.begin_y $first.end_x $first.end_y $second.begin_x $second.begin_y $second.end_x $second.end_y $endpointClearance
            if ($null -ne $crossing) {
                & $addFinding "connector-crossing" "hard" @([string]$first.name, [string]$second.name) "Two unrelated straight arrow routes cross near ($([math]::Round($crossing.x,2)),$([math]::Round($crossing.y,2)))." "Assign separate routing lanes, change connection sites, or replace one route with a clean elbow connector." "Unrelated connector routes do not cross."
            }
        }
    }

    $seriesGroups = @{}
    foreach ($shape in $summaries) {
        $name = [string]$shape.name
        $key = [regex]::Replace($name, '[-_]\d+$', '-*')
        if ($key -eq $name) { continue }
        if (-not $seriesGroups.ContainsKey($key)) { $seriesGroups[$key] = @() }
        $seriesGroups[$key] += $shape
    }
    foreach ($entry in $seriesGroups.GetEnumerator()) {
        $group = @($entry.Value)
        if ($group.Count -lt 3) { continue }
        $centersX = @($group | ForEach-Object { [double]$_.left + [double]$_.width / 2 })
        $centersY = @($group | ForEach-Object { [double]$_.top + [double]$_.height / 2 })
        $rangeX = ($centersX | Measure-Object -Maximum).Maximum - ($centersX | Measure-Object -Minimum).Minimum
        $rangeY = ($centersY | Measure-Object -Maximum).Maximum - ($centersY | Measure-Object -Minimum).Minimum
        $verticalSeries = $rangeY -ge $rangeX
        $crossValues = if ($verticalSeries) { @($group | ForEach-Object { [double]$_.left }) } else { @($group | ForEach-Object { [double]$_.top }) }
        $crossSpread = ($crossValues | Measure-Object -Maximum).Maximum - ($crossValues | Measure-Object -Minimum).Minimum
        if ($crossSpread -gt $alignmentTolerance) {
            & $addFinding "repeated-series-misalignment" "hard" @($group | ForEach-Object { [string]$_.name }) "Repeated series '$($entry.Key)' has cross-axis spread $([math]::Round($crossSpread,2)) pt, above $alignmentTolerance pt." "Use powerpoint_align_shapes on the repeated objects, preserving the reference anchor." "Cross-axis edges are equal within $alignmentTolerance pt."
        }
        $ordered = if ($verticalSeries) { @($group | Sort-Object top) } else { @($group | Sort-Object left) }
        $gaps = @()
        for ($index = 1; $index -lt $ordered.Count; $index += 1) {
            $previous = $ordered[$index - 1]
            $current = $ordered[$index]
            $gaps += if ($verticalSeries) { [double]$current.top - ([double]$previous.top + [double]$previous.height) } else { [double]$current.left - ([double]$previous.left + [double]$previous.width) }
        }
        if ($gaps.Count -ge 2) {
            $gapSpread = ($gaps | Measure-Object -Maximum).Maximum - ($gaps | Measure-Object -Minimum).Minimum
            if ($gapSpread -gt (2 * $alignmentTolerance)) {
                & $addFinding "repeated-series-unequal-spacing" "warning" @($group | ForEach-Object { [string]$_.name }) "Repeated series '$($entry.Key)' gap spread is $([math]::Round($gapSpread,2)) pt." "Use powerpoint_distribute_shapes or exact coordinates after fixing endpoints." "Repeated gaps differ by no more than $([math]::Round(2 * $alignmentTolerance,2)) pt."
            }
        }
    }

    $hardFailures = @($findings | Where-Object { $_.severity -eq "hard" })
    $warnings = @($findings | Where-Object { $_.severity -eq "warning" })
    $findingsTruncated = $findings.Count -ge $maxFindings
    return [ordered]@{
        backend = "powerpoint"
        slide_index = $slideIndex
        slide_size = [ordered]@{ width = $slideWidth; height = $slideHeight }
        object_counts = [ordered]@{
            total = [int]$slide.Shapes.Count
            pictures = @($summaries | Where-Object { [int]$_.type -in @(11, 13) }).Count
            native_or_composite = [int]$slide.Shapes.Count - @($summaries | Where-Object { [int]$_.type -in @(11, 13) }).Count
        }
        picture_audit = $pictureAudits
        hard_failure_count = $hardFailures.Count
        warning_count = $warnings.Count
        findings_truncated = $findingsTruncated
        passed = $hardFailures.Count -eq 0 -and -not $findingsTruncated
        findings = @($findings)
        reviewer_contract = "Every hard finding must be converted by the Corrector into named geometry or decomposition operations, executed by the Drawer, rerendered, and audited again."
    }
}

function Invoke-ActivateSlide {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $index = [int](Get-Argument $Arguments "slide_index")
    $null = Get-Slide $presentation $index
    Show-Slide $application $index $true
    return [ordered]@{ slide_index = $index; activated = $true }
}

function Invoke-AddSlide {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $position = [int](Get-Argument $Arguments "position" ($presentation.Slides.Count + 1))
    if ($position -lt 1 -or $position -gt ($presentation.Slides.Count + 1)) {
        throw "position $position is outside the valid insertion range 1..$($presentation.Slides.Count + 1)."
    }
    $layoutMap = @{ blank = 12; title = 1; text = 2 }
    $layoutName = [string](Get-Argument $Arguments "layout" "blank")
    $slide = $presentation.Slides.Add($position, $layoutMap[$layoutName])
    $name = Get-Argument $Arguments "name"
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $slide.Name = [string]$name }
    Show-Slide $application $position
    return [ordered]@{ slide_index = [int]$position; slide_id = [int]$slide.SlideID; name = [string]$slide.Name; layout = $layoutName }
}

function Invoke-AddTextbox {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    if (-not (Test-Property $Arguments "text_autofit")) { $Arguments | Add-Member -NotePropertyName text_autofit -NotePropertyValue "none" }
    if (-not (Test-Property $Arguments "word_wrap")) { $Arguments | Add-Member -NotePropertyName word_wrap -NotePropertyValue $true }
    $shape = $slide.Shapes.AddTextbox(
        1,
        [single](Get-Argument $Arguments "left"),
        [single](Get-Argument $Arguments "top"),
        [single](Get-Argument $Arguments "width"),
        [single](Get-Argument $Arguments "height")
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    Set-ShapeFromArguments $shape $Arguments
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Resolve-AutoShapeType {
    param($Arguments)
    if (Test-Property $Arguments "shape_type_id") {
        return [int](Get-Argument $Arguments "shape_type_id")
    }
    $requested = [string](Get-Argument $Arguments "shape")
    if ([string]::IsNullOrWhiteSpace($requested)) { throw "Provide shape or shape_type_id." }
    $legacyAliases = @{
        triangle = "msoShapeIsoscelesTriangle"
        ellipse = "msoShapeOval"
        rounded_rectangle = "msoShapeRoundedRectangle"
        can = "msoShapeCan"
    }
    $candidate = $requested.Trim()
    if ($legacyAliases.ContainsKey($candidate.ToLowerInvariant())) { $candidate = $legacyAliases[$candidate.ToLowerInvariant()] }
    $interop = Import-OfficeInteropMetadata
    if (-not $interop.office_core) { throw "Office enum metadata is unavailable; cannot resolve AutoShape '$requested'." }
    $enumType = [Microsoft.Office.Core.MsoAutoShapeType]
    foreach ($enumName in [Enum]::GetNames($enumType)) {
        $pluginName = Convert-EnumNameToPluginName $enumName "msoShape"
        if ($enumName -ieq $candidate -or $pluginName -ieq $candidate) {
            return [int][Enum]::Parse($enumType, $enumName)
        }
    }
    throw "Unknown AutoShape '$requested'. Call powerpoint_get_capabilities and use an auto_shapes.plugin_name, office_name, or value."
}

function Invoke-AddShape {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    if (Test-Property $Arguments "text") {
        if (-not (Test-Property $Arguments "text_autofit")) { $Arguments | Add-Member -NotePropertyName text_autofit -NotePropertyValue "none" }
        if (-not (Test-Property $Arguments "word_wrap")) { $Arguments | Add-Member -NotePropertyName word_wrap -NotePropertyValue $true }
    }
    $shapeType = Resolve-AutoShapeType $Arguments
    $shape = $slide.Shapes.AddShape(
        $shapeType,
        [single](Get-Argument $Arguments "left"),
        [single](Get-Argument $Arguments "top"),
        [single](Get-Argument $Arguments "width"),
        [single](Get-Argument $Arguments "height")
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    Set-ShapeFromArguments $shape $Arguments
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Invoke-AddImage {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $imagePath = [IO.Path]::GetFullPath([string](Get-Argument $Arguments "image_path"))
    if (-not (Test-Path -LiteralPath $imagePath -PathType Leaf)) { throw "Image file does not exist: $imagePath" }
    $rasterReason = [string](Get-Argument $Arguments "raster_reason")
    if ([string]::IsNullOrWhiteSpace($rasterReason) -or $rasterReason.Trim().Length -lt 8) {
        throw "raster_reason must specifically explain why this exact region cannot be recreated with native editable PowerPoint objects."
    }
    $atomicRasterUnit = [bool](Get-Argument $Arguments "atomic_raster_unit" $false)
    if (-not $atomicRasterUnit) {
        throw "atomic_raster_unit=true is required. Split grids, comparisons, montages, stacks, panels, and other multi-part visuals into separate picture objects first."
    }
    $containsReconstructableContent = [bool](Get-Argument $Arguments "contains_reconstructable_content" $true)
    if ($containsReconstructableContent) {
        throw "contains_reconstructable_content must be false. Rebuild text, borders, arrows, legends, axes, tables, and regular plots as native PowerPoint objects."
    }
    $decompositionNote = [string](Get-Argument $Arguments "decomposition_note" "")
    if ([string]::IsNullOrWhiteSpace($decompositionNote) -or $decompositionNote.Trim().Length -lt 8) {
        throw "decomposition_note must state what was separated and rebuilt natively, or why this picture cannot be split further."
    }
    $cropProperties = @(
        "crop_left_percent", "crop_top_percent", "crop_right_percent", "crop_bottom_percent",
        "crop_left_points", "crop_top_points", "crop_right_points", "crop_bottom_points"
    )
    $hasCrop = $false
    foreach ($cropProperty in $cropProperties) {
        if (Test-Property $Arguments $cropProperty) { $hasCrop = $true; break }
    }
    $sourceIsTight = [bool](Get-Argument $Arguments "source_is_tightly_cropped" $false)
    if (-not $sourceIsTight -and -not $hasCrop) {
        throw "source_is_tightly_cropped=false requires at least one crop_* field. Crop away all surrounding reconstructable content before insertion."
    }
    $horizontalCropPercent = [double](Get-Argument $Arguments "crop_left_percent" 0) + [double](Get-Argument $Arguments "crop_right_percent" 0)
    $verticalCropPercent = [double](Get-Argument $Arguments "crop_top_percent" 0) + [double](Get-Argument $Arguments "crop_bottom_percent" 0)
    if ($horizontalCropPercent -ge 100 -or $verticalCropPercent -ge 100) { throw "Opposing crop percentages must total less than 100%." }
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $shape = $slide.Shapes.AddPicture(
        $imagePath,
        0,
        -1,
        [single](Get-Argument $Arguments "left"),
        [single](Get-Argument $Arguments "top"),
        [single](Get-Argument $Arguments "width"),
        [single](Get-Argument $Arguments "height")
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    if ([bool](Get-Argument $Arguments "lock_aspect_ratio" $false)) { $shape.LockAspectRatio = -1 }
    $baseWidth = [double]$shape.Width
    $baseHeight = [double]$shape.Height
    $cropMap = @{
        crop_left_percent = "CropLeft"; crop_top_percent = "CropTop"; crop_right_percent = "CropRight"; crop_bottom_percent = "CropBottom"
        crop_left_points = "CropLeft"; crop_top_points = "CropTop"; crop_right_points = "CropRight"; crop_bottom_points = "CropBottom"
    }
    foreach ($cropProperty in $cropMap.Keys) {
        if (-not (Test-Property $Arguments $cropProperty)) { continue }
        $cropValue = [double](Get-Argument $Arguments $cropProperty)
        if ($cropProperty.EndsWith("_percent")) {
            $cropValue = if ($cropProperty -match "left|right") { $baseWidth * $cropValue / 100.0 } else { $baseHeight * $cropValue / 100.0 }
        }
        $pictureProperty = $cropMap[$cropProperty]
        $shape.PictureFormat.$pictureProperty = [single]$cropValue
    }
    $altText = if (Test-Property $Arguments "alt_text") { [string](Get-Argument $Arguments "alt_text") } else { "Raster-only visual evidence: $rasterReason" }
    $shape.AlternativeText = $altText
    try {
        $shape.Tags.Add("ScientificIllustratorRasterReason", $rasterReason)
        $shape.Tags.Add("ScientificIllustratorSourceTightlyCropped", [string]$sourceIsTight)
        $shape.Tags.Add("ScientificIllustratorAtomicRasterUnit", [string]$atomicRasterUnit)
        $shape.Tags.Add("ScientificIllustratorContainsReconstructableContent", [string]$containsReconstructableContent)
        $shape.Tags.Add("ScientificIllustratorDecompositionNote", $decompositionNote)
        $shape.Tags.Add("ScientificIllustratorSourcePath", $imagePath)
    }
    catch {}
    Show-Slide $application $slideIndex
    $summary = New-ShapeSummary $shape
    $summary.raster_reason = $rasterReason
    $summary.source_is_tightly_cropped = $sourceIsTight
    $summary.atomic_raster_unit = $atomicRasterUnit
    $summary.contains_reconstructable_content = $containsReconstructableContent
    $summary.decomposition_note = $decompositionNote
    return $summary
}

function Get-ArrowStyle {
    param([string]$Name)
    $map = @{ none = 1; open = 2; triangle = 3; stealth = 4; diamond = 5; oval = 6 }
    return $map[$Name]
}

function Invoke-AddLine {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $beginX = [double](Get-Argument $Arguments "begin_x")
    $beginY = [double](Get-Argument $Arguments "begin_y")
    $endX = [double](Get-Argument $Arguments "end_x")
    $endY = [double](Get-Argument $Arguments "end_y")
    $startClearance = [double](Get-Argument $Arguments "start_clearance" 0)
    $endClearance = [double](Get-Argument $Arguments "end_clearance" 0)
    $deltaX = $endX - $beginX
    $deltaY = $endY - $beginY
    $length = [math]::Sqrt($deltaX * $deltaX + $deltaY * $deltaY)
    if ($length -le 0.0001) { throw "Line length must be greater than zero." }
    if (($startClearance + $endClearance) -ge $length) { throw "start_clearance + end_clearance must be less than the line length." }
    $unitX = $deltaX / $length
    $unitY = $deltaY / $length
    $actualBeginX = $beginX + $unitX * $startClearance
    $actualBeginY = $beginY + $unitY * $startClearance
    $actualEndX = $endX - $unitX * $endClearance
    $actualEndY = $endY - $unitY * $endClearance
    $shape = $slide.Shapes.AddLine(
        [single]$actualBeginX,
        [single]$actualBeginY,
        [single]$actualEndX,
        [single]$actualEndY
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    Set-ShapeAppearance $shape $Arguments
    $shape.Line.BeginArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "start_arrow" "none"))
    $shape.Line.EndArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "end_arrow" "none"))
    try {
        $shape.Tags.Add("ScientificIllustratorBeginX", [string]$actualBeginX)
        $shape.Tags.Add("ScientificIllustratorBeginY", [string]$actualBeginY)
        $shape.Tags.Add("ScientificIllustratorEndX", [string]$actualEndX)
        $shape.Tags.Add("ScientificIllustratorEndY", [string]$actualEndY)
        $shape.Tags.Add("ScientificIllustratorStartClearance", [string]$startClearance)
        $shape.Tags.Add("ScientificIllustratorEndClearance", [string]$endClearance)
    }
    catch {}
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Invoke-AddConnector {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $sourceArgs = [pscustomobject]@{ shape_name = [string](Get-Argument $Arguments "source_name") }
    $targetArgs = [pscustomobject]@{ shape_name = [string](Get-Argument $Arguments "target_name") }
    $source = Find-Shape $slide $sourceArgs
    $target = Find-Shape $slide $targetArgs
    $sourceSite = [int](Get-Argument $Arguments "source_site" 1)
    $targetSite = [int](Get-Argument $Arguments "target_site" 1)
    if ($sourceSite -gt $source.ConnectionSiteCount) { throw "source_site $sourceSite exceeds '$($source.Name)' connection-site count $($source.ConnectionSiteCount)." }
    if ($targetSite -gt $target.ConnectionSiteCount) { throw "target_site $targetSite exceeds '$($target.Name)' connection-site count $($target.ConnectionSiteCount)." }
    $connectorTypeMap = @{ straight = 1; elbow = 2; curve = 3 }
    $connectorType = [string](Get-Argument $Arguments "connector_type" "elbow")
    $shape = $slide.Shapes.AddConnector($connectorTypeMap[$connectorType], 0, 0, 100, 100)
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    $shape.ConnectorFormat.BeginConnect($source, $sourceSite)
    $shape.ConnectorFormat.EndConnect($target, $targetSite)
    $shape.RerouteConnections()
    $shape.ZOrder(1)
    Set-ShapeAppearance $shape $Arguments
    $shape.Line.BeginArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "start_arrow" "none"))
    $shape.Line.EndArrowheadStyle = Get-ArrowStyle ([string](Get-Argument $Arguments "end_arrow" "triangle"))
    try {
        $shape.Tags.Add("ScientificIllustratorSourceName", [string]$source.Name)
        $shape.Tags.Add("ScientificIllustratorTargetName", [string]$target.Name)
        $shape.Tags.Add("ScientificIllustratorSourceSite", [string]$sourceSite)
        $shape.Tags.Add("ScientificIllustratorTargetSite", [string]$targetSite)
    }
    catch {}
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Set-TableCellStyle {
    param($Cell, $Arguments)
    $cellShape = $Cell.Shape
    if (Test-Property $Arguments "fill_color") {
        $cellShape.Fill.Visible = -1
        $cellShape.Fill.ForeColor.RGB = Convert-HexToOfficeRgb (Get-Argument $Arguments "fill_color")
        $cellShape.Fill.Solid()
    }
    if (Test-Property $Arguments "fill_transparency") {
        $cellShape.Fill.Visible = -1
        $cellShape.Fill.Transparency = [single]([double](Get-Argument $Arguments "fill_transparency") / 100.0)
    }
    Set-ShapeText $cellShape $Arguments
    if (Test-Property $Arguments "cell_margin") {
        $margin = [single](Get-Argument $Arguments "cell_margin")
        $cellShape.TextFrame.MarginLeft = $margin
        $cellShape.TextFrame.MarginRight = $margin
        $cellShape.TextFrame.MarginTop = $margin
        $cellShape.TextFrame.MarginBottom = $margin
    }
    if ((Test-Property $Arguments "border_color") -or (Test-Property $Arguments "border_width")) {
        foreach ($borderIndex in 1..4) {
            $border = $Cell.Borders.Item($borderIndex)
            if (Test-Property $Arguments "border_color") {
                $border.Visible = -1
                $border.ForeColor.RGB = Convert-HexToOfficeRgb (Get-Argument $Arguments "border_color")
            }
            if (Test-Property $Arguments "border_width") {
                $borderWidth = [double](Get-Argument $Arguments "border_width")
                if ($borderWidth -eq 0) { $border.Visible = 0 }
                else { $border.Visible = -1; $border.Weight = [single]$borderWidth }
            }
        }
    }
}

function Invoke-AddTable {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $rows = [int](Get-Argument $Arguments "rows")
    $columns = [int](Get-Argument $Arguments "columns")
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $data = @()
    if (Test-Property $Arguments "data") { $data = @(Get-Argument $Arguments "data") }
    if ($data.Count -gt $rows) { throw "Table data has $($data.Count) rows but rows=$rows." }
    for ($rowIndex = 0; $rowIndex -lt $data.Count; $rowIndex += 1) {
        if (@($data[$rowIndex]).Count -gt $columns) { throw "Table data row $($rowIndex + 1) has more than columns=$columns values." }
    }
    $shape = $slide.Shapes.AddTable(
        $rows,
        $columns,
        [single](Get-Argument $Arguments "left"),
        [single](Get-Argument $Arguments "top"),
        [single](Get-Argument $Arguments "width"),
        [single](Get-Argument $Arguments "height")
    )
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
    $headerRows = [math]::Min($rows, [int](Get-Argument $Arguments "header_rows" 1))
    $sharedTextProperties = @("font_name", "font_size", "font_color", "bold", "italic", "alignment", "vertical_alignment", "cell_margin", "fill_transparency", "border_color", "border_width")
    for ($row = 1; $row -le $rows; $row += 1) {
        for ($column = 1; $column -le $columns; $column += 1) {
            $cellArguments = [ordered]@{}
            foreach ($propertyName in $sharedTextProperties) {
                if (Test-Property $Arguments $propertyName) { $cellArguments[$propertyName] = Get-Argument $Arguments $propertyName }
            }
            if (Test-Property $Arguments "fill_color") { $cellArguments.fill_color = Get-Argument $Arguments "fill_color" }
            if ($row -le $headerRows) {
                if (Test-Property $Arguments "header_fill_color") { $cellArguments.fill_color = Get-Argument $Arguments "header_fill_color" }
                if (Test-Property $Arguments "header_font_color") { $cellArguments.font_color = Get-Argument $Arguments "header_font_color" }
                $cellArguments.bold = [bool](Get-Argument $Arguments "header_bold" $true)
            }
            elseif ([bool](Get-Argument $Arguments "banded_rows" $false) -and (($row - $headerRows) % 2 -eq 1) -and (Test-Property $Arguments "band_fill_color")) {
                $cellArguments.fill_color = Get-Argument $Arguments "band_fill_color"
            }
            if ($row -le $data.Count -and $column -le @($data[$row - 1]).Count) {
                $value = @($data[$row - 1])[$column - 1]
                $cellArguments.text = if ($null -eq $value) { "" } else { [string]$value }
            }
            Set-TableCellStyle $shape.Table.Cell($row, $column) ([pscustomobject]$cellArguments)
        }
    }
    if (Test-Property $Arguments "cell_styles") {
        foreach ($cellStyle in @(Get-Argument $Arguments "cell_styles")) {
            $row = [int](Get-Argument $cellStyle "row")
            $column = [int](Get-Argument $cellStyle "column")
            if ($row -lt 1 -or $row -gt $rows -or $column -lt 1 -or $column -gt $columns) {
                throw "cell_styles entry ($row,$column) is outside the table bounds $rows x $columns."
            }
            Set-TableCellStyle $shape.Table.Cell($row, $column) $cellStyle
        }
    }
    Show-Slide $application $slideIndex
    $summary = New-ShapeSummary $shape
    $summary.rows = $rows
    $summary.columns = $columns
    return $summary
}

function Invoke-UpdateTableCell {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    if ($shape.HasTable -ne -1) { throw "Shape '$($shape.Name)' is not a native PowerPoint table." }
    $row = [int](Get-Argument $Arguments "row")
    $column = [int](Get-Argument $Arguments "column")
    $rows = [int]$shape.Table.Rows.Count
    $columns = [int]$shape.Table.Columns.Count
    if ($row -lt 1 -or $row -gt $rows -or $column -lt 1 -or $column -gt $columns) {
        throw "Cell ($row,$column) is outside table '$($shape.Name)' bounds $rows x $columns."
    }
    $cell = $shape.Table.Cell($row, $column)
    Set-TableCellStyle $cell $Arguments
    Show-Slide $application $slideIndex
    return [ordered]@{
        slide_index = $slideIndex
        table_name = [string]$shape.Name
        row = $row
        column = $column
        text = Get-ShapeText $cell.Shape
        updated = $true
    }
}

function Invoke-UpdateTableLayout {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    if ($shape.HasTable -ne -1) { throw "Shape '$($shape.Name)' is not a native PowerPoint table." }
    $updated = $false
    if (Test-Property $Arguments "column_widths") {
        $widths = @([object[]](Get-Argument $Arguments "column_widths"))
        if ($widths.Count -ne [int]$shape.Table.Columns.Count) { throw "column_widths count $($widths.Count) does not match table column count $($shape.Table.Columns.Count)." }
        for ($index = 1; $index -le $widths.Count; $index += 1) { $shape.Table.Columns.Item($index).Width = [single]$widths[$index - 1] }
        $updated = $true
    }
    if (Test-Property $Arguments "row_heights") {
        $heights = @([object[]](Get-Argument $Arguments "row_heights"))
        if ($heights.Count -ne [int]$shape.Table.Rows.Count) { throw "row_heights count $($heights.Count) does not match table row count $($shape.Table.Rows.Count)." }
        for ($index = 1; $index -le $heights.Count; $index += 1) { $shape.Table.Rows.Item($index).Height = [single]$heights[$index - 1] }
        $updated = $true
    }
    if (-not $updated) { throw "Provide column_widths and/or row_heights." }
    Show-Slide $application $slideIndex
    $summary = New-ShapeSummary $shape
    $summary.column_widths = @(1..$shape.Table.Columns.Count | ForEach-Object { [math]::Round([double]$shape.Table.Columns.Item($_).Width, 2) })
    $summary.row_heights = @(1..$shape.Table.Rows.Count | ForEach-Object { [math]::Round([double]$shape.Table.Rows.Item($_).Height, 2) })
    return $summary
}

function Resolve-ChartType {
    param($Arguments)
    if (Test-Property $Arguments "chart_type_id") { return [int](Get-Argument $Arguments "chart_type_id") }
    $requested = [string](Get-Argument $Arguments "chart_type")
    if ([string]::IsNullOrWhiteSpace($requested)) { throw "Provide chart_type or chart_type_id." }
    $interop = Import-OfficeInteropMetadata
    if (-not $interop.excel_chart_types) { throw "Excel chart type metadata is unavailable; cannot resolve chart type '$requested'." }
    $enumType = [Microsoft.Office.Interop.Excel.XlChartType]
    foreach ($enumName in [Enum]::GetNames($enumType)) {
        $pluginName = Convert-EnumNameToPluginName $enumName "xl"
        if ($enumName -ieq $requested -or $pluginName -ieq $requested) {
            return [int][Enum]::Parse($enumType, $enumName)
        }
    }
    throw "Unknown chart type '$requested'. Call powerpoint_get_capabilities and use a chart_types.plugin_name, office_name, or value."
}

function Invoke-AddChart {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $categories = @(Get-Argument $Arguments "categories")
    $seriesItems = @(Get-Argument $Arguments "series")
    foreach ($series in $seriesItems) {
        $values = @(Get-Argument $series "values")
        if ($values.Count -ne $categories.Count) {
            throw "Chart series '$([string](Get-Argument $series 'name'))' has $($values.Count) values but there are $($categories.Count) categories."
        }
    }
    $chartType = Resolve-ChartType $Arguments
    $shape = $null
    $workbook = $null
    try {
        try {
            $shape = $slide.Shapes.AddChart2(
                -1,
                $chartType,
                [single](Get-Argument $Arguments "left"),
                [single](Get-Argument $Arguments "top"),
                [single](Get-Argument $Arguments "width"),
                [single](Get-Argument $Arguments "height"),
                $true
            )
        }
        catch {
            $shape = $slide.Shapes.AddChart(
                $chartType,
                [single](Get-Argument $Arguments "left"),
                [single](Get-Argument $Arguments "top"),
                [single](Get-Argument $Arguments "width"),
                [single](Get-Argument $Arguments "height")
            )
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $shape.Name = [string]$name }
        $chart = $shape.Chart
        $chartData = $chart.ChartData
        $chartData.Activate()
        for ($attempt = 0; $attempt -lt 30 -and $null -eq $workbook; $attempt += 1) {
            try { $workbook = $chartData.Workbook } catch {}
            if ($null -eq $workbook) { Start-Sleep -Milliseconds 200 }
        }
        if ($null -eq $workbook) { throw "PowerPoint created the native chart but did not expose its embedded data workbook after 6 seconds." }
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Cells.Clear()
        $worksheet.Cells.Item(1, 1).Value2 = ""
        for ($seriesIndex = 0; $seriesIndex -lt $seriesItems.Count; $seriesIndex += 1) {
            $series = $seriesItems[$seriesIndex]
            $worksheet.Cells.Item(1, $seriesIndex + 2).Value2 = [string](Get-Argument $series "name")
        }
        for ($categoryIndex = 0; $categoryIndex -lt $categories.Count; $categoryIndex += 1) {
            $worksheet.Cells.Item($categoryIndex + 2, 1).Value2 = [string]$categories[$categoryIndex]
            for ($seriesIndex = 0; $seriesIndex -lt $seriesItems.Count; $seriesIndex += 1) {
                $values = @(Get-Argument $seriesItems[$seriesIndex] "values")
                $worksheet.Cells.Item($categoryIndex + 2, $seriesIndex + 2).Value2 = [double]$values[$categoryIndex]
            }
        }
        $lastCell = $worksheet.Cells.Item($categories.Count + 1, $seriesItems.Count + 1)
        $sourceAddress = "='$($worksheet.Name.Replace("'", "''"))'!`$A`$1:$($lastCell.Address($true, $true, 1, $false))"
        $chart.SetSourceData($sourceAddress)
        if (Test-Property $Arguments "title") {
            $chart.HasTitle = -1
            $chart.ChartTitle.Text = [string](Get-Argument $Arguments "title")
        }
        else { $chart.HasTitle = 0 }
        $chart.HasLegend = if ([bool](Get-Argument $Arguments "has_legend" $true)) { -1 } else { 0 }
        if ($chart.HasLegend -eq -1) {
            $legendMap = @{ right = -4152; left = -4131; top = -4160; bottom = -4107 }
            $chart.Legend.Position = $legendMap[[string](Get-Argument $Arguments "legend_position" "right")]
        }
        if (Test-Property $Arguments "chart_style") { $chart.ChartStyle = [int](Get-Argument $Arguments "chart_style") }
        if (Test-Property $Arguments "category_axis_title") {
            $axis = $chart.Axes(1, 1)
            $axis.HasTitle = -1
            $axis.AxisTitle.Text = [string](Get-Argument $Arguments "category_axis_title")
        }
        if (Test-Property $Arguments "value_axis_title") {
            $axis = $chart.Axes(2, 1)
            $axis.HasTitle = -1
            $axis.AxisTitle.Text = [string](Get-Argument $Arguments "value_axis_title")
        }
        try { $workbook.Close($true) } catch {}
        $workbook = $null
        Show-Slide $application $slideIndex
        $summary = New-ShapeSummary $shape
        $summary.chart_type_id = $chartType
        $summary.category_count = $categories.Count
        $summary.series_count = $seriesItems.Count
        return $summary
    }
    catch {
        try { if ($null -ne $workbook) { $workbook.Close($false) } } catch {}
        try { if ($null -ne $shape) { $shape.Delete() } } catch {}
        throw
    }
}

function Invoke-DuplicateShape {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $source = Find-Shape $slide $Arguments
    $newName = [string](Get-Argument $Arguments "new_name")
    Assert-ShapeNameAvailable $slide $newName
    $shape = $source.Duplicate().Item(1)
    $shape.Name = $newName
    Set-ShapeGeometry $shape $Arguments
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Invoke-GroupShapes {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shapeNames = @([object[]](Get-Argument $Arguments "shape_names"))
    foreach ($shapeName in $shapeNames) {
        $null = Find-Shape $slide ([pscustomobject]@{ shape_name = [string]$shapeName })
    }
    $name = Get-Argument $Arguments "name"
    Assert-ShapeNameAvailable $slide ([string]$name)
    $group = $slide.Shapes.Range([object[]]$shapeNames).Group()
    if (-not [string]::IsNullOrWhiteSpace([string]$name)) { $group.Name = [string]$name }
    Show-Slide $application $slideIndex
    return New-ShapeSummary $group
}

function Invoke-UngroupShape {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    if ([int]$shape.Type -ne 6) { throw "Shape '$($shape.Name)' is not a native group." }
    $range = $shape.Ungroup()
    $members = @()
    for ($index = 1; $index -le $range.Count; $index += 1) { $members += [pscustomobject](New-ShapeSummary $range.Item($index)) }
    Show-Slide $application $slideIndex
    return [ordered]@{ slide_index = $slideIndex; ungrouped = $true; member_count = $members.Count; members = $members }
}

function Invoke-SetZOrder {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    $commandMap = @{ bring_to_front = 0; send_to_back = 1; bring_forward = 2; send_backward = 3 }
    $command = [string](Get-Argument $Arguments "command")
    $repeat = [int](Get-Argument $Arguments "repeat" 1)
    foreach ($step in 1..$repeat) { $shape.ZOrder($commandMap[$command]) }
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Invoke-AlignShapes {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shapeNames = @([object[]](Get-Argument $Arguments "shape_names"))
    foreach ($shapeName in $shapeNames) { $null = Find-Shape $slide ([pscustomobject]@{ shape_name = [string]$shapeName }) }
    $alignmentMap = @{ left = 0; center = 1; right = 2; top = 3; middle = 4; bottom = 5 }
    $alignment = [string](Get-Argument $Arguments "alignment")
    $relativeToSlide = [string](Get-Argument $Arguments "relative_to" "selection") -eq "slide"
    $range = $slide.Shapes.Range([object[]]$shapeNames)
    $range.Align($alignmentMap[$alignment], $(if ($relativeToSlide) { -1 } else { 0 }))
    $items = @()
    foreach ($shapeName in $shapeNames) { $items += [pscustomobject](New-ShapeSummary (Find-Shape $slide ([pscustomobject]@{ shape_name = [string]$shapeName }))) }
    Show-Slide $application $slideIndex
    return [ordered]@{ slide_index = $slideIndex; alignment = $alignment; relative_to = if ($relativeToSlide) { "slide" } else { "selection" }; shapes = $items }
}

function Invoke-DistributeShapes {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shapeNames = @([object[]](Get-Argument $Arguments "shape_names"))
    foreach ($shapeName in $shapeNames) { $null = Find-Shape $slide ([pscustomobject]@{ shape_name = [string]$shapeName }) }
    $directionMap = @{ horizontal = 0; vertical = 1 }
    $direction = [string](Get-Argument $Arguments "direction")
    $relativeToSlide = [string](Get-Argument $Arguments "relative_to" "selection") -eq "slide"
    $range = $slide.Shapes.Range([object[]]$shapeNames)
    $range.Distribute($directionMap[$direction], $(if ($relativeToSlide) { -1 } else { 0 }))
    $items = @()
    foreach ($shapeName in $shapeNames) { $items += [pscustomobject](New-ShapeSummary (Find-Shape $slide ([pscustomobject]@{ shape_name = [string]$shapeName }))) }
    Show-Slide $application $slideIndex
    return [ordered]@{ slide_index = $slideIndex; direction = $direction; relative_to = if ($relativeToSlide) { "slide" } else { "selection" }; shapes = $items }
}

function Invoke-UpdateShape {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    if (Test-Property $Arguments "new_name") {
        $newName = [string](Get-Argument $Arguments "new_name")
        Assert-ShapeNameAvailable $slide $newName $shape
        $shape.Name = $newName
    }
    Set-ShapeFromArguments $shape $Arguments
    Show-Slide $application $slideIndex
    return New-ShapeSummary $shape
}

function Invoke-DeleteShape {
    param($Arguments)
    if (-not [bool](Get-Argument $Arguments "confirm" $false)) { throw "confirm=true is required to delete a shape." }
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $shape = Find-Shape $slide $Arguments
    $summary = New-ShapeSummary $shape
    $shape.Delete()
    Show-Slide $application $slideIndex
    return [ordered]@{ deleted = $true; slide_index = $slideIndex; shape = $summary }
}

function Invoke-ExportSlideImage {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $slideIndex = [int](Get-Argument $Arguments "slide_index")
    $slide = Get-Slide $presentation $slideIndex
    $outputPath = [IO.Path]::GetFullPath([string](Get-Argument $Arguments "output_path"))
    $extension = [IO.Path]::GetExtension($outputPath).ToLowerInvariant()
    if ($extension -notin @(".png", ".jpg", ".jpeg")) { throw "output_path must end with .png, .jpg, or .jpeg." }
    if ((Test-Path -LiteralPath $outputPath) -and -not [bool](Get-Argument $Arguments "overwrite" $false)) {
        throw "Output exists; pass overwrite=true: $outputPath"
    }
    $directory = [IO.Path]::GetDirectoryName($outputPath)
    if (-not [string]::IsNullOrWhiteSpace($directory)) { $null = New-Item -ItemType Directory -Force -Path $directory }
    if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }
    $format = if ($extension -eq ".png") { "PNG" } else { "JPG" }
    $requestedWidth = [int](Get-Argument $Arguments "width" 1920)
    $heightWasExplicit = Test-Property $Arguments "height"
    $slideWidth = [double]$presentation.PageSetup.SlideWidth
    $slideHeight = [double]$presentation.PageSetup.SlideHeight
    if ($slideWidth -le 0 -or $slideHeight -le 0) { throw "PowerPoint returned invalid slide dimensions for image export." }
    $requestedHeight = if ($heightWasExplicit) { [int](Get-Argument $Arguments "height") } else { [int][math]::Round($requestedWidth * $slideHeight / $slideWidth) }
    $preserveAspectRatio = [bool](Get-Argument $Arguments "preserve_aspect_ratio" $true)
    $width = $requestedWidth
    $height = $requestedHeight
    if ($preserveAspectRatio -and $heightWasExplicit) {
        $scale = [math]::Min($requestedWidth / $slideWidth, $requestedHeight / $slideHeight)
        $width = [math]::Max(1, [int][math]::Round($slideWidth * $scale))
        $height = [math]::Max(1, [int][math]::Round($slideHeight * $scale))
    }
    $slide.Export($outputPath, $format, $width, $height)
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw "PowerPoint did not create the requested image: $outputPath" }
    $file = Get-Item -LiteralPath $outputPath
    return [ordered]@{
        output_path = $outputPath
        slide_index = $slideIndex
        requested_width = $requestedWidth
        requested_height = $requestedHeight
        width = $width
        height = $height
        aspect_ratio_preserved = $preserveAspectRatio
        bytes = [long]$file.Length
        mime_type = if ($format -eq "PNG") { "image/png" } else { "image/jpeg" }
    }
}

function Invoke-Save {
    param($Arguments)
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $outputPath = Get-Argument $Arguments "output_path"
    if ([string]::IsNullOrWhiteSpace([string]$outputPath)) {
        if ([string]::IsNullOrWhiteSpace([string]$presentation.Path)) {
            throw "The active presentation has never been saved; provide an absolute output_path."
        }
        $presentation.Save()
        return [ordered]@{ saved = $true; output_path = [string]$presentation.FullName; format = "in-place" }
    }
    $resolved = [IO.Path]::GetFullPath([string]$outputPath)
    $format = [string](Get-Argument $Arguments "format")
    if ([string]::IsNullOrWhiteSpace($format)) {
        $format = switch ([IO.Path]::GetExtension($resolved).ToLowerInvariant()) {
            ".pptx" { "pptx" }
            ".pdf" { "pdf" }
            default { throw "output_path must end with .pptx or .pdf, or provide format explicitly." }
        }
    }
    $expectedExtension = if ($format -eq "pdf") { ".pdf" } else { ".pptx" }
    if ([IO.Path]::GetExtension($resolved).ToLowerInvariant() -ne $expectedExtension) {
        throw "format '$format' requires an $expectedExtension output_path."
    }
    if ((Test-Path -LiteralPath $resolved) -and -not [bool](Get-Argument $Arguments "overwrite" $false)) {
        throw "Output exists; pass overwrite=true: $resolved"
    }
    $directory = [IO.Path]::GetDirectoryName($resolved)
    if (-not [string]::IsNullOrWhiteSpace($directory)) { $null = New-Item -ItemType Directory -Force -Path $directory }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }
    if ($format -eq "pdf") {
        $presentation.ExportAsFixedFormat($resolved, 2)
    }
    else {
        $presentation.SaveCopyAs($resolved, 24)
    }
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw "PowerPoint did not create the requested file: $resolved" }
    $file = Get-Item -LiteralPath $resolved
    return [ordered]@{ saved = $true; output_path = $resolved; format = $format; bytes = [long]$file.Length }
}

function Invoke-ClosePresentation {
    param($Arguments)
    if (-not [bool](Get-Argument $Arguments "confirm" $false)) { throw "confirm=true is required to close the active presentation." }
    $application = Get-PowerPointApplication $false
    $presentation = Get-ActivePresentation $application
    $name = [string]$presentation.Name
    $saveChanges = [string](Get-Argument $Arguments "save_changes" "discard")
    $savedPath = ""
    if ($saveChanges -eq "save") {
        $outputPath = Get-Argument $Arguments "output_path"
        if (-not [string]::IsNullOrWhiteSpace([string]$outputPath)) {
            $resolved = [IO.Path]::GetFullPath([string]$outputPath)
            if ([IO.Path]::GetExtension($resolved).ToLowerInvariant() -ne ".pptx") { throw "output_path must end with .pptx." }
            if ((Test-Path -LiteralPath $resolved) -and -not [bool](Get-Argument $Arguments "overwrite" $false)) {
                throw "Output exists; pass overwrite=true: $resolved"
            }
            $directory = [IO.Path]::GetDirectoryName($resolved)
            if (-not [string]::IsNullOrWhiteSpace($directory)) { $null = New-Item -ItemType Directory -Force -Path $directory }
            if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }
            $presentation.SaveAs($resolved, 24)
            $savedPath = $resolved
        }
        else {
            if ([string]::IsNullOrWhiteSpace([string]$presentation.Path)) { throw "The presentation has no file path; provide output_path when save_changes=save." }
            $presentation.Save()
            $savedPath = [string]$presentation.FullName
        }
    }
    else {
        $presentation.Saved = -1
    }
    $presentation.Close()
    return [ordered]@{ closed = $true; presentation_name = $name; save_changes = $saveChanges; output_path = $savedPath }
}

function Invoke-QuitApplication {
    param($Arguments)
    if (-not [bool](Get-Argument $Arguments "confirm" $false)) { throw "confirm=true is required to quit PowerPoint." }
    $application = Get-PowerPointApplication $false
    if ($null -eq $application) { throw "PowerPoint is not running." }
    $presentationCount = [int]$application.Presentations.Count
    if ($presentationCount -ne 0) { throw "Refusing to quit PowerPoint because it has $presentationCount open presentation(s). Close only the intended presentation first." }
    $actualProcessId = Get-PowerPointProcessId $application
    $expectedProcessId = [int](Get-Argument $Arguments "expected_process_id")
    if ($actualProcessId -ne $expectedProcessId) { throw "Refusing to quit PowerPoint: expected process id $expectedProcessId but the active COM application is process $actualProcessId." }
    $application.Quit()
    return [ordered]@{ quit = $true; process_id = $actualProcessId; presentation_count = 0 }
}

function Invoke-Action {
    param([string]$Action, $Arguments)
    switch ($Action) {
        "status" { return Invoke-Status }
        "capabilities" { return Invoke-Capabilities $Arguments }
        "launch" { return Invoke-Launch $Arguments }
        "new_presentation" { return Invoke-NewPresentation $Arguments }
        "inspect" { return Invoke-Inspect $Arguments }
        "audit_figure" { return Invoke-AuditFigure $Arguments }
        "activate_slide" { return Invoke-ActivateSlide $Arguments }
        "add_slide" { return Invoke-AddSlide $Arguments }
        "add_textbox" { return Invoke-AddTextbox $Arguments }
        "add_shape" { return Invoke-AddShape $Arguments }
        "add_image" { return Invoke-AddImage $Arguments }
        "add_line" { return Invoke-AddLine $Arguments }
        "add_connector" { return Invoke-AddConnector $Arguments }
        "add_table" { return Invoke-AddTable $Arguments }
        "update_table_cell" { return Invoke-UpdateTableCell $Arguments }
        "update_table_layout" { return Invoke-UpdateTableLayout $Arguments }
        "add_chart" { return Invoke-AddChart $Arguments }
        "duplicate_shape" { return Invoke-DuplicateShape $Arguments }
        "group_shapes" { return Invoke-GroupShapes $Arguments }
        "ungroup_shape" { return Invoke-UngroupShape $Arguments }
        "set_z_order" { return Invoke-SetZOrder $Arguments }
        "align_shapes" { return Invoke-AlignShapes $Arguments }
        "distribute_shapes" { return Invoke-DistributeShapes $Arguments }
        "update_shape" { return Invoke-UpdateShape $Arguments }
        "delete_shape" { return Invoke-DeleteShape $Arguments }
        "export_slide_image" { return Invoke-ExportSlideImage $Arguments }
        "save" { return Invoke-Save $Arguments }
        "close_presentation" { return Invoke-ClosePresentation $Arguments }
        "quit_application" { return Invoke-QuitApplication $Arguments }
        default { throw "Unknown PowerPoint bridge action: $Action" }
    }
}

try {
    $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($PayloadBase64))
    $payload = $json | ConvertFrom-Json
    $requestedFocusPolicy = Get-Argument $payload.arguments "focus_policy" $env:SCIENTIFIC_ILLUSTRATOR_FOCUS_POLICY
    $script:FocusPolicy = Normalize-FocusPolicy $requestedFocusPolicy
    $previousForegroundWindow = if ($script:FocusPolicy -eq "preserve") { Get-ForegroundWindowHandle } else { [IntPtr]::Zero }
    try {
        $result = Invoke-Action ([string]$payload.action) $payload.arguments
    }
    finally {
        if ($script:FocusPolicy -eq "preserve" -and [string]$payload.action -ne "activate_slide") {
            Restore-ForegroundWindow $previousForegroundWindow
        }
    }
    [Console]::Out.Write(($result | ConvertTo-Json -Depth 12 -Compress))
}
catch {
    $details = @(
        [string]$_.Exception.Message
        [string]$_.InvocationInfo.PositionMessage
        [string]$_.ScriptStackTrace
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    [Console]::Error.Write(($details -join [Environment]::NewLine))
    exit 1
}
