import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ImageDimensions {
    width: number;
    height: number;
}

export class CompressResizeImage implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    
    // UI Elements
    private _mainContainer: HTMLDivElement;
    private _uploadButton: HTMLButtonElement;
    private _thumbnailContainer: HTMLDivElement;
    private _thumbnailImage: HTMLImageElement;
    private _removeButton: HTMLButtonElement;
    private _loadingIndicator: HTMLDivElement;
    
    // State
    private _imageData: string | null = null;
    private _isProcessing = false;

    constructor() {
        // Empty
    }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        // Create main container
        this._mainContainer = document.createElement("div");
        this._mainContainer.className = "compress-resize-image-container";

        // Create upload trigger (styled as minimal, subtle area with + icon)
        this._uploadButton = document.createElement("button");
        this._uploadButton.className = "upload-button";
        this._uploadButton.setAttribute("type", "button");
        this._uploadButton.innerHTML = `
            <span class="upload-icon" aria-hidden="true">＋</span>
            <span class="button-text">Upload Image</span>
        `;
        this._uploadButton.addEventListener("click", this.onUploadClick.bind(this));

        // Create thumbnail container (initially hidden)
        this._thumbnailContainer = document.createElement("div");
        this._thumbnailContainer.className = "thumbnail-container";
        this._thumbnailContainer.style.display = "none";

        // Create thumbnail image
        this._thumbnailImage = document.createElement("img");
        this._thumbnailImage.className = "thumbnail-image";
        this._thumbnailImage.alt = "Uploaded image";

        // Create remove button
        this._removeButton = document.createElement("button");
        this._removeButton.className = "remove-button";
        this._removeButton.innerHTML = "✕";
        this._removeButton.title = "Remove image";
        this._removeButton.addEventListener("click", this.onRemoveClick.bind(this));

        // Create loading indicator
        this._loadingIndicator = document.createElement("div");
        this._loadingIndicator.style.display = "none";
        this._loadingIndicator.innerHTML = `
            <div class="loading-spinner"></div>
            <div class="loading-text">Processing image...</div>
        `;

        // Append elements
        this._thumbnailContainer.appendChild(this._thumbnailImage);
        this._thumbnailContainer.appendChild(this._removeButton);
        
        this._mainContainer.appendChild(this._uploadButton);
        this._mainContainer.appendChild(this._thumbnailContainer);
        this._mainContainer.appendChild(this._loadingIndicator);
        
        this._container.appendChild(this._mainContainer);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;

        // Update button text if provided
        const buttonText = context.parameters.ButtonText.raw || "Upload Image";
        const buttonTextElement = this._uploadButton.querySelector(".button-text");
        if (buttonTextElement) {
            buttonTextElement.textContent = buttonText;
        }

        // Sync UI with bound ImageData (handle both set and cleared states)
        const incoming = context.parameters.ImageData.raw ?? null;
        if (incoming !== this._imageData) {
            this._imageData = incoming;
            if (this._imageData && this._imageData.length > 0) {
                this.displayThumbnail(this._imageData);
            } else {
                // Cleared externally or empty -> reset UI
                this._thumbnailImage.src = "";
                this._thumbnailContainer.style.display = "none";
                this._uploadButton.style.display = "flex";
            }
        }
    }

    private async onUploadClick(): Promise<void> {
        if (this._isProcessing) return;

        try {
            const captureMode = this._context.parameters.CaptureMode.raw || "0";
            let imageFile: ComponentFramework.FileObject;

            if (captureMode === "1") {
                // Camera mode
                const maxHeight = this._context.parameters.MaxHeight.raw;
                const maxWidth = this._context.parameters.MaxWidth.raw;
                
                const captureOptions: ComponentFramework.DeviceApi.CaptureImageOptions = {
                    allowEdit: false,
                    preferFrontCamera: false,
                    height: maxHeight && maxHeight > 0 ? maxHeight : 480,
                    width: maxWidth && maxWidth > 0 ? maxWidth : 640,
                    quality: 100 // We'll compress later
                };
                imageFile = await this._context.device.captureImage(captureOptions);
            } else {
                // File picker mode (environment)
                const pickOptions: ComponentFramework.DeviceApi.PickFileOptions = {
                    accept: "image",
                    allowMultipleFiles: false,
                    maximumAllowedFileSize: 10485760 // 10MB
                };
                const files = await this._context.device.pickFile(pickOptions);
                if (!files || files.length === 0) return;
                imageFile = files[0];
            }

            // Show loading indicator
            this.showLoading(true);

            // Process the image
            await this.processImage(imageFile);

        } catch (error) {
            console.error("Error uploading image:", error);
            this.showError("Failed to upload image. Please try again.");
        } finally {
            this.showLoading(false);
        }
    }

    private async processImage(imageFile: ComponentFramework.FileObject): Promise<void> {
        try {
            // FileObject.fileContent is already a base64 string, not an ArrayBuffer
            const base64Data = imageFile.fileContent;
            const mimeType = imageFile.mimeType || "image/jpeg";
            
            // Ensure proper data URL format
            let fullBase64: string;
            if (base64Data.startsWith('data:')) {
                fullBase64 = base64Data;
            } else {
                fullBase64 = `data:${mimeType};base64,${base64Data}`;
            }

            // Load image to get dimensions
            const img = await this.loadImage(fullBase64);
            
            // Get compression settings
            const maxWidth = this._context.parameters.MaxWidth.raw || 0;
            const maxHeight = this._context.parameters.MaxHeight.raw || 0;
            const compressionMode = this._context.parameters.CompressionMode.raw || "0";
            const quality = this._context.parameters.Quality.raw || 80;
            const targetSizeKB = this._context.parameters.TargetSizeKB.raw || 100;

            // Calculate new dimensions if resizing is needed
            const newDimensions = this.calculateDimensions(img.width, img.height, maxWidth, maxHeight);

            // Compress the image
            let compressedBase64: string;
            if (compressionMode === "1") {
                // Target size mode
                compressedBase64 = await this.compressToTargetSize(img, newDimensions, targetSizeKB, mimeType);
            } else {
                // Quality mode
                compressedBase64 = await this.compressWithQuality(img, newDimensions, quality / 100, mimeType);
            }

            // Store the compressed image
            this._imageData = compressedBase64;
            this._notifyOutputChanged();

            // Display thumbnail
            this.displayThumbnail(compressedBase64);

        } catch (error) {
            console.error("Error processing image:", error);
            throw error;
        }
    }

    private calculateDimensions(
        originalWidth: number,
        originalHeight: number,
        maxWidth: number,
        maxHeight: number
    ): ImageDimensions {
        // If no max dimensions are set, keep original size
        if (maxWidth === 0 && maxHeight === 0) {
            return { width: originalWidth, height: originalHeight };
        }

        let width = originalWidth;
        let height = originalHeight;

        // Calculate aspect ratio
        const aspectRatio = originalWidth / originalHeight;

        // Resize based on max dimensions while maintaining aspect ratio
        if (maxWidth > 0 && width > maxWidth) {
            width = maxWidth;
            height = width / aspectRatio;
        }

        if (maxHeight > 0 && height > maxHeight) {
            height = maxHeight;
            width = height * aspectRatio;
        }

        return { width: Math.round(width), height: Math.round(height) };
    }

    private async compressWithQuality(
        img: HTMLImageElement,
        dimensions: ImageDimensions,
        quality: number,
        mimeType: string
    ): Promise<string> {
        const canvas = document.createElement("canvas");
        canvas.width = dimensions.width;
        canvas.height = dimensions.height;

        const ctx = canvas.getContext("2d");
        if (!ctx) throw new Error("Could not get canvas context");

        // Draw image on canvas with new dimensions
        ctx.drawImage(img, 0, 0, dimensions.width, dimensions.height);

        // Convert to base64 with specified quality
        return canvas.toDataURL(mimeType, quality);
    }

    private async compressToTargetSize(
        img: HTMLImageElement,
        dimensions: ImageDimensions,
        targetSizeKB: number,
        mimeType: string
    ): Promise<string> {
        let quality = 0.9;
        let compressed = "";
        let currentSizeKB = 0;
        const minQuality = 0.1;
        const maxIterations = 10;
        let iteration = 0;

        // Iteratively reduce quality until we reach target size
        while (iteration < maxIterations) {
            compressed = await this.compressWithQuality(img, dimensions, quality, mimeType);
            currentSizeKB = this.getBase64SizeKB(compressed);

            if (currentSizeKB <= targetSizeKB || quality <= minQuality) {
                break;
            }

            // Adjust quality based on how far we are from target
            const ratio = targetSizeKB / currentSizeKB;
            quality = Math.max(minQuality, quality * ratio * 0.9);
            iteration++;
        }

        return compressed;
    }

    private getBase64SizeKB(base64String: string): number {
        // Remove data URL prefix if present
        const base64Data = base64String.split(',')[1] || base64String;
        // Calculate size: each base64 character is 6 bits, so divide by 8 to get bytes, then by 1024 for KB
        const sizeInBytes = (base64Data.length * 3) / 4;
        return sizeInBytes / 1024;
    }

    private loadImage(src: string): Promise<HTMLImageElement> {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = reject;
            img.src = src;
        });
    }

    private displayThumbnail(base64Data: string): void {
        this._thumbnailImage.src = base64Data;
        this._thumbnailContainer.style.display = "flex";
        this._uploadButton.style.display = "none";
    }

    private onRemoveClick(): void {
        // Clear the image data and propagate an explicit empty string to bound output
        this._imageData = ""; // empty string signals clearing the bound field
        this._thumbnailImage.src = "";
        this._thumbnailContainer.style.display = "none";
        this._uploadButton.style.display = "flex";
        this._notifyOutputChanged();
    }

    private showLoading(show: boolean): void {
        this._isProcessing = show;
        this._loadingIndicator.style.display = show ? "flex" : "none";
        this._uploadButton.style.display = show ? "none" : (this._imageData ? "none" : "flex");
        this._thumbnailContainer.style.display = show ? "none" : (this._imageData ? "flex" : "none");
    }

    private showError(message: string): void {
        const errorDiv = document.createElement("div");
        errorDiv.className = "error-message";
        errorDiv.textContent = message;
        this._mainContainer.appendChild(errorDiv);

        // Remove error after 5 seconds
        setTimeout(() => {
            if (errorDiv.parentNode === this._mainContainer) {
                this._mainContainer.removeChild(errorDiv);
            }
        }, 5000);
    }

    public getOutputs(): IOutputs {
        return {
            // When _imageData is null -> no change. When empty string -> explicitly clear. When base64 string -> set value.
            ImageData: this._imageData !== null ? this._imageData : undefined
        };
    }

    public destroy(): void {
        // Clean up event listeners
        if (this._uploadButton) {
            this._uploadButton.removeEventListener("click", this.onUploadClick.bind(this));
        }
        if (this._removeButton) {
            this._removeButton.removeEventListener("click", this.onRemoveClick.bind(this));
        }
    }
}
