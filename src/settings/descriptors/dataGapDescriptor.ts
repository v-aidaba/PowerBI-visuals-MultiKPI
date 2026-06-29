/**
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import ToggleSwitch = formattingSettings.ToggleSwitch;
import FormattingSettingsSlice = formattingSettings.Slice;
import TextInput = formattingSettings.TextInput;
import ColorPicker = formattingSettings.ColorPicker;

import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;

import { BaseDescriptor } from "./baseDescriptor";

export class DataGapDescriptor extends BaseDescriptor {
    public name: string = "dataGap";
    public displayNameKey: string = "Visual_DataGap";
    public descriptionKey: string = "Visual_DataGapDescription";

    public defaultColorValue: string = "#ffeb3b";
    public defaultBackgroundValue: string = "";

    public gapMessage: TextInput = new TextInput({
        name: "gapMessage",
        displayNameKey: "Visual_DataGapMessage",
        descriptionKey: "Visual_DataGapMessageDescription",
        value: "⚠️ ${1} missing days detected",
        placeholder: "Enter gap message template"
    });

    public backgroundColor: ColorPicker = new ColorPicker({
        name: "backgroundColor",
        displayNameKey: "Visual_BackgroundColor",
        value: {value: this.defaultBackgroundValue}
    });

    public color: ColorPicker = new ColorPicker({
        name: "color",
        displayNameKey: "Visual_Color",
        value: { value: this.defaultColorValue }
    });

    public slices: FormattingSettingsSlice[] = [
        this.gapMessage,
        this.backgroundColor,
        this.color
    ];

    topLevelSlice: ToggleSwitch = this.isShown;

    public processHighContrastMode(colorPalette: ISandboxExtendedColorPalette): void {
        const isHighContrast: boolean = colorPalette.isHighContrast;

        this.color.visible = isHighContrast ? false : this.color.visible;
        this.color.value = isHighContrast ? colorPalette.foreground : this.color.value;

        this.backgroundColor.visible = isHighContrast ? false : this.backgroundColor.visible;
        this.backgroundColor.value = isHighContrast ? colorPalette.foreground : this.backgroundColor.value;
    }
}