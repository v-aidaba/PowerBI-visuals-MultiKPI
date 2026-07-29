/*
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

import "../styles/styles.less";

import powerbi from "powerbi-visuals-api";

import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import { dispatch, Dispatch } from "d3-dispatch";
import { select as d3Select } from "d3-selection";

import { DataConverter } from "./converter/data/dataConverter";

import { EventName } from "./event/eventName";

import { Settings } from "./settings/settings";

import { DataOrderConverter } from "./converter/data/dataOrderConverter";
import { DataOrderConverterWithDuplicates } from "./converter/data/dataOrderConverterWithDuplicates";

import { IDataRepresentation } from "./converter/data/dataRepresentation";
import { isValidDate } from "./utils/isValidDate";

import { RootComponent } from "./visualComponent/rootComponent";
import { IVisualComponent } from "./visualComponent/visualComponent";
import { IVisualComponentRenderOptions } from "./visualComponent/visualComponentRenderOptions";

import { ScaleService } from "./services/scaleService";

// powerbi.extensibility.utils.tooltip
import { ITooltipServiceWrapper, TooltipServiceWrapper } from "powerbi-visuals-utils-tooltiputils";

export class MultiKpi implements powerbi.extensibility.visual.IVisual {
    private dataConverter: DataConverter;

    private minViewport: powerbi.IViewport = {
        height: 95,
        width: 200,
    };

    private dataRepresentation: IDataRepresentation;
    private localizationManager: ILocalizationManager;
    private settings: Settings;
    private formattingSettingsService: FormattingSettingsService;
    private viewport: powerbi.IViewport;
    private eventDispatcher: Dispatch<object> = dispatch(...Object.keys(EventName));
    private tooltipServiceWrapper: ITooltipServiceWrapper;
    private host: powerbi.extensibility.visual.IVisualHost;
    private selectionManager: ISelectionManager;
    private element: HTMLElement;
    private isLandingPageOn: boolean = false;
    private landingPageRemoved: boolean = false;
    private landingPage: HTMLElement | null = null;

    public rootComponent: IVisualComponent<IVisualComponentRenderOptions>;

    constructor(options: powerbi.extensibility.visual.VisualConstructorOptions) {
        const {
            element,
            host,
        } = options;

        this.host = host;

        this.element = element;

        this.localizationManager = options.host.createLocalizationManager();
        this.formattingSettingsService = new FormattingSettingsService(this.localizationManager);

        this.tooltipServiceWrapper = new TooltipServiceWrapper(
            {
                handleTouchDelay: 0,
                rootElement: options.element,
                tooltipService: host.tooltipService,
            });

        this.dataConverter = new DataConverter({
            createSelectionIdBuilder: host.createSelectionIdBuilder.bind(host),
        });

        this.eventDispatcher.on(
            EventName.onChartChange,
            this.onChartChange.bind(this),
        );

        this.eventDispatcher.on(
            EventName.onChartViewChange,
            this.onChartViewChange.bind(this),
        );

        this.eventDispatcher.on(
            EventName.onChartViewReset,
            () => this.render(this.dataRepresentation, this.settings, this.viewport),
        );

        this.rootComponent = new RootComponent({
            element: d3Select(element),
            eventDispatcher: this.eventDispatcher,
            scaleService: new ScaleService(element),
            colorPalette: host.colorPalette,
            tooltipServiceWrapper: this.tooltipServiceWrapper,
            localizationManager: this.localizationManager,
        });

        this.selectionManager = this.host.createSelectionManager();

        const visualSelection = d3Select(element);
        visualSelection.on("contextmenu", (event: PointerEvent, dataPoint) => {
            this.selectionManager.showContextMenu(dataPoint ? dataPoint : {}, {
                x: event.clientX,
                y: event.clientY
            });
            event.preventDefault();
        });
    }

    public update(options: powerbi.extensibility.visual.VisualUpdateOptions) {
        if (!this.dataConverter || !this.rootComponent) {
            return;
        }

        try {
            this.host.eventService.renderingStarted(options);

            if (this.handleLandingPage(options)) {
                this.settings = this.formattingSettingsService.populateFormattingSettingsModel(Settings, options?.dataViews?.[0]);
                this.settings.parse(this.host.colorPalette, this.localizationManager);
                this.rootComponent.hide?.();
                this.renderNoDataMessage(null);
            } else {
                const dataView: powerbi.DataView = options?.dataViews?.[0];

                this.viewport = this.getViewport(options?.viewport);

                this.settings = this.formattingSettingsService.populateFormattingSettingsModel(Settings, dataView);

                this.dataRepresentation = this.dataConverter.convert({
                    dataView,
                    settings: this.settings,
                    viewport: this.viewport,
                });

                this.settings.parse(this.host.colorPalette, this.localizationManager);

                const noDataMessage: string = this.getNoDataMessage(dataView, this.dataRepresentation);

                if (noDataMessage) {
                    this.rootComponent.hide?.();
                    this.renderNoDataMessage(noDataMessage);
                } else {
                    this.renderNoDataMessage(null);
                    this.rootComponent.show?.();

                    this.render(
                        this.dataRepresentation,
                        this.settings,
                        this.viewport,
                    );
                }
            }
        } catch (ex) {
            this.host.eventService.renderingFailed(options, JSON.stringify(ex));
        }
        this.host.eventService.renderingFinished(options);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.settings);
    }    

    private getNoDataMessage(dataView: powerbi.DataView, data: IDataRepresentation): string | null {
        const hasDate: boolean = !!dataView?.categorical?.categories?.[0]?.values?.length;
        const hasValues: boolean = !!dataView?.categorical?.values?.length;

        if (!hasDate && !hasValues) {
            return this.localizationManager.getDisplayName("Visual_EmptyState_AddDateAndValues");
        }

        if (!hasDate) {
            return this.localizationManager.getDisplayName("Visual_EmptyState_AddDate");
        }

        if (!hasValues) {
            return this.localizationManager.getDisplayName("Visual_EmptyState_AddValues");
        }

        if (!this.hasValidDate(dataView)) {
            return this.localizationManager.getDisplayName("Visual_EmptyState_InvalidDate");
        }

        if (!data?.series?.length || !this.hasValidData(data)) {
            return this.localizationManager.getDisplayName("Visual_EmptyState_NoValidData");
        }

        return null;
    }

    private hasValidData(data: IDataRepresentation): boolean {
        return data.series.some((series) =>
            series?.points?.some((point) => point != null && !isNaN(point.y)),
        );
    }

    private hasValidDate(dataView: powerbi.DataView): boolean {
        const categories: powerbi.DataViewCategoryColumn[] | undefined = dataView?.categorical?.categories;

        if (!categories?.length) {
            return false;
        }

        const dateCategory: powerbi.DataViewCategoryColumn =
            categories.find((category) => category?.source?.roles?.dateColumn) ?? categories[0];

        return !!dateCategory?.values?.some(
            (value) => isValidDate(value),
        );
    }

    private handleLandingPage(options: powerbi.extensibility.visual.VisualUpdateOptions): boolean {
        const hasData: boolean = !!options?.dataViews?.[0]?.metadata?.columns?.length;

        if (!hasData) {
            if (!this.isLandingPageOn) {
                this.isLandingPageOn = true;
                this.landingPageRemoved = false;
                this.landingPage = this.createLandingPage();
                this.element.appendChild(this.landingPage);
            }

            return true;
        }

        if (this.isLandingPageOn && !this.landingPageRemoved) {
            this.isLandingPageOn = false;
            this.landingPageRemoved = true;

            if (this.landingPage) {
                this.landingPage.remove();
                this.landingPage = null;
            }
        }

        return false;
    }

    private createLandingPage(): HTMLElement {
        const container: HTMLElement = document.createElement("div");
        container.classList.add("multiKpi_landingPage");

        const icon: HTMLElement = document.createElement("div");
        icon.classList.add("multiKpi_landingPage_icon");
        container.appendChild(icon);

        const title: HTMLElement = document.createElement("div");
        title.classList.add("multiKpi_landingPage_title");
        title.textContent = this.localizationManager.getDisplayName("Visual_LandingPage_Title");
        container.appendChild(title);

        const description: HTMLElement = document.createElement("div");
        description.classList.add("multiKpi_landingPage_description");
        description.textContent = this.localizationManager.getDisplayName("Visual_LandingPage_Description");
        container.appendChild(description);

        return container;
    }

    private renderNoDataMessage(message: string | null): void {
        d3Select(this.element)
            .selectAll("div.multiKpi_emptyState")
            .data(message ? [message] : [])
            .join("div")
            .classed("multiKpi_emptyState", true)
            .text((text: string) => text);
    }

    private render(
        data: IDataRepresentation,
        settings: Settings,
        viewport: powerbi.IViewport,
    ): void {
        this.rootComponent.render({
            data,
            settings,
            viewport,
        });
    }

    private getViewport(currentViewport: powerbi.IViewport): powerbi.IViewport {
        if (!currentViewport) {
            return { ...this.minViewport };
        }

        return {
            height: Math.max(this.minViewport.height, currentViewport.height),
            width: Math.max(this.minViewport.width, currentViewport.width),
        };
    }

    private onChartChange(name: string): void {
        const dataOrderConverter: DataOrderConverter = new DataOrderConverter();

        this.dataRepresentation = dataOrderConverter.convert({
            data: this.dataRepresentation,
            firstSeriesName: name,
        });

        this.render(
            this.dataRepresentation,
            this.settings,
            this.viewport,
        );
    }

    private onChartViewChange(name: string): void {
        const dataOrderConverter: DataOrderConverterWithDuplicates = new DataOrderConverterWithDuplicates();

        const dataRepresentation: IDataRepresentation = dataOrderConverter.convert({
            data: this.dataRepresentation,
            firstSeriesName: name,
        });

        this.render(
            dataRepresentation,
            this.settings,
            this.viewport,
        );
    }
}
