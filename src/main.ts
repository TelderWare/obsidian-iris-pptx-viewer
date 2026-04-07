import { Plugin, FileView, WorkspaceLeaf, TFile, Notice, setIcon } from "obsidian";
import { PDFDocument } from "pdf-lib";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import * as pdfjsWorker from "pdfjs-dist/legacy/build/pdf.worker.mjs";

(globalThis as any).pdfjsWorker = pdfjsWorker;
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";

const CONVERTER_VIEW = "pptx-converter";
const SLIDESHOW_VIEW = "pdf-slideshow";
const PPTX_EXT = "pptx";
const SLIDESHOW_MARKER = "obsidian-pptx-slideshow";

// ── LibreOffice detection ──

function findLibreOffice(): string | null {
	const { existsSync } = require("fs");
	const candidates: string[] = [];

	if (process.platform === "win32") {
		for (const base of [
			process.env["ProgramFiles"] || "C:\\Program Files",
			process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)",
			process.env["LOCALAPPDATA"] || "",
		]) {
			if (base) candidates.push(`${base}\\LibreOffice\\program\\soffice.exe`);
		}
	} else if (process.platform === "darwin") {
		candidates.push("/Applications/LibreOffice.app/Contents/MacOS/soffice");
		candidates.push("/opt/homebrew/bin/soffice");
		candidates.push("/usr/local/bin/soffice");
	} else {
		candidates.push("/usr/bin/soffice");
		candidates.push("/usr/bin/libreoffice");
		candidates.push("/snap/bin/libreoffice");
		candidates.push("/usr/local/bin/soffice");
	}

	for (const p of candidates) {
		if (existsSync(p)) return p;
	}

	try {
		const { execSync } = require("child_process");
		const cmd = process.platform === "win32" ? "where soffice" : "which soffice";
		const result = execSync(cmd, { encoding: "utf-8" }).trim().split("\n")[0];
		if (result && existsSync(result)) return result;
	} catch { /* not found */ }

	return null;
}

// ── PPTX processing ──

async function stripAnimations(pptxPath: string, tmpDir: string): Promise<string> {
	const { readFile, writeFile } = require("fs/promises");
	const path = require("path");
	const JSZip = require("jszip");

	const data = await readFile(pptxPath);
	const zip = await JSZip.loadAsync(data);

	const slideFiles = Object.keys(zip.files).filter(
		(name: string) => /^ppt\/slides\/slide\d+\.xml$/.test(name)
	);

	for (const slidePath of slideFiles) {
		let xml: string = await zip.file(slidePath).async("string");
		xml = xml.replace(/<p:timing[\s\S]*?<\/p:timing>/g, "");
		zip.file(slidePath, xml);
	}

	const cleanedPath = path.join(tmpDir, "cleaned.pptx");
	const buf = await zip.generateAsync({ type: "nodebuffer" });
	await writeFile(cleanedPath, buf);
	return cleanedPath;
}

async function convertToPdf(pptxPath: string, sofficePath: string): Promise<Buffer> {
	const { execFile } = require("child_process");
	const { readFile, mkdtemp, rm } = require("fs/promises");
	const path = require("path");
	const os = require("os");

	const tmpDir = await mkdtemp(path.join(os.tmpdir(), "pptx-viewer-"));

	try {
		const cleanedPath = await stripAnimations(pptxPath, tmpDir);

		const { stdout, stderr } = await new Promise<{ stdout: string; stderr: string }>((resolve, reject) => {
			execFile(
				sofficePath,
				["--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmpDir, cleanedPath],
				{ timeout: 30000 },
				(err: Error | null, stdout: string, stderr: string) => {
					if (err) reject(err);
					else resolve({ stdout, stderr });
				}
			);
		});

		const pdfPath = path.join(tmpDir, "cleaned.pdf");
		const { existsSync } = require("fs");
		if (!existsSync(pdfPath)) {
			const { readdirSync } = require("fs");
			const files = readdirSync(tmpDir);
			throw new Error(
				`LibreOffice did not produce a PDF.\n` +
				`stdout: ${stdout}\nstderr: ${stderr}\n` +
				`Files in tmpDir: ${files.join(", ")}`
			);
		}

		const result = await readFile(pdfPath);
		if (result.length === 0) {
			throw new Error("LibreOffice produced an empty PDF.");
		}

		// Stamp the PDF with slideshow marker
		const pdfDoc = await PDFDocument.load(result);
		pdfDoc.setSubject(SLIDESHOW_MARKER);
		const stamped = await pdfDoc.save();
		return Buffer.from(stamped);
	} finally {
		await rm(tmpDir, { recursive: true, force: true }).catch(() => {});
	}
}

async function hasSlideshowMarker(file: TFile, vault: any): Promise<boolean> {
	try {
		const data = await vault.readBinary(file);
		const pdfDoc = await PDFDocument.load(data, { updateMetadata: false });
		return pdfDoc.getSubject() === SLIDESHOW_MARKER;
	} catch {
		return false;
	}
}

// ── PPTX Converter View (transient) ──

class PptxConverterView extends FileView {
	private plugin: PptxViewerPlugin;

	constructor(leaf: WorkspaceLeaf, plugin: PptxViewerPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType() { return CONVERTER_VIEW; }
	getDisplayText() { return this.file?.basename ?? "PPTX"; }
	getIcon() { return "projector"; }
	canAcceptExtension(ext: string) { return ext === PPTX_EXT; }

	async onLoadFile(file: TFile) {
		const { contentEl } = this;
		contentEl.empty();
		const loading = contentEl.createDiv({ cls: "pptx-loading" });
		loading.createDiv({ cls: "pptx-loading-spinner" });
		loading.createSpan({ text: "Converting to PDF\u2026" });
		await this.plugin.convertAndReplace(file, this.leaf);
	}

	async onUnloadFile() {
		this.contentEl.empty();
	}
}

// ── PDF Slideshow View ──

class PdfSlideshowView extends FileView {
	private pdfDoc: PDFDocumentProxy | null = null;
	private currentSlide = 0;
	private slideCount = 0;
	private counterEl: HTMLElement | null = null;
	private canvas: HTMLCanvasElement | null = null;
	private wrapper: HTMLElement | null = null;
	private resizeObserver: ResizeObserver | null = null;
	private renderGeneration = 0;
	private activeRenderTask: any = null;

	getViewType() { return SLIDESHOW_VIEW; }
	getDisplayText() { return this.file?.basename ?? "Slideshow"; }
	getIcon() { return "projector"; }
	canAcceptExtension(ext: string) { return ext === "pdf"; }

	async onLoadFile(file: TFile) {
		const { contentEl } = this;
		contentEl.empty();

		try {
			const data = await this.app.vault.readBinary(file);
			const loadingTask = pdfjsLib.getDocument({
				data,
				useWorkerFetch: false,
				isEvalSupported: false,
				useSystemFonts: true,
			});
			this.pdfDoc = await loadingTask.promise;
			this.slideCount = this.pdfDoc.numPages;
			this.currentSlide = 0;

			this.buildUI(contentEl);
			await this.renderSlide();

			this.resizeObserver = new ResizeObserver(() => {
				this.renderSlide();
			});
			if (this.wrapper) {
				this.resizeObserver.observe(this.wrapper);
			}
		} catch (e) {
			contentEl.empty();
			const err = contentEl.createDiv({ cls: "pptx-error" });
			err.createSpan({ text: "Failed to load presentation" });
			new Notice(`Slideshow error: ${String(e)}`, 0);
			console.error("PDF Slideshow error:", e);
		}
	}

	async onUnloadFile() {
		if (this.resizeObserver) {
			this.resizeObserver.disconnect();
			this.resizeObserver = null;
		}
		if (this.pdfDoc) {
			this.pdfDoc.destroy();
			this.pdfDoc = null;
		}
		this.slideCount = 0;
		this.contentEl.empty();
	}

	private buildUI(root: HTMLElement) {
		const container = root.createDiv({ cls: "pptx-viewer-container" });

		const toolbar = container.createDiv({ cls: "pptx-toolbar" });
		const prevBtn = toolbar.createDiv({ cls: "pptx-nav-btn" });
		setIcon(prevBtn, "chevron-left");
		prevBtn.addEventListener("click", (e) => {
			e.stopPropagation();
			this.goToSlide(this.currentSlide - 1);
		});
		this.counterEl = toolbar.createDiv({ cls: "pptx-slide-counter" });
		const nextBtn = toolbar.createDiv({ cls: "pptx-nav-btn" });
		setIcon(nextBtn, "chevron-right");
		nextBtn.addEventListener("click", (e) => {
			e.stopPropagation();
			this.goToSlide(this.currentSlide + 1);
		});

		this.wrapper = container.createDiv({ cls: "pptx-slide-wrapper" });
		this.canvas = this.wrapper.createEl("canvas", { cls: "pptx-slide-canvas" });

		container.tabIndex = 0;
		container.addEventListener("click", () => {
			this.goToSlide(this.currentSlide + 1);
		});
		container.addEventListener("keydown", (e: KeyboardEvent) => {
			if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
				e.preventDefault();
				this.goToSlide(this.currentSlide - 1);
			} else if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === " ") {
				e.preventDefault();
				this.goToSlide(this.currentSlide + 1);
			}
		});
		container.focus();
	}

	private async goToSlide(index: number) {
		if (index < 0 || index >= this.slideCount) return;
		this.currentSlide = index;
		await this.renderSlide();
	}

	private async renderSlide() {
		if (!this.pdfDoc || !this.canvas || !this.counterEl || !this.wrapper) return;

		if (this.activeRenderTask) {
			this.activeRenderTask.cancel();
			this.activeRenderTask = null;
		}

		const gen = ++this.renderGeneration;

		this.counterEl.textContent = `${this.currentSlide + 1} / ${this.slideCount}`;

		const page: PDFPageProxy = await this.pdfDoc.getPage(this.currentSlide + 1);
		if (gen !== this.renderGeneration) return;

		const wrapperRect = this.wrapper.getBoundingClientRect();
		const availW = wrapperRect.width;
		const availH = wrapperRect.height;
		if (availW <= 0 || availH <= 0) return;

		const unscaledViewport = page.getViewport({ scale: 1 });
		const scaleW = availW / unscaledViewport.width;
		const scaleH = availH / unscaledViewport.height;
		const displayScale = Math.min(scaleW, scaleH);

		const dpr = window.devicePixelRatio || 1;
		const renderScale = displayScale * dpr;
		const viewport = page.getViewport({ scale: renderScale });

		this.canvas.width = viewport.width;
		this.canvas.height = viewport.height;
		this.canvas.style.width = `${viewport.width / dpr}px`;
		this.canvas.style.height = `${viewport.height / dpr}px`;

		const ctx = this.canvas.getContext("2d")!;
		ctx.clearRect(0, 0, viewport.width, viewport.height);

		const renderTask = page.render({
			canvasContext: ctx,
			viewport: viewport,
		});
		this.activeRenderTask = renderTask;

		try {
			await renderTask.promise;
		} catch (e: any) {
			if (e?.name === "RenderingCancelledException") return;
			throw e;
		}

		this.activeRenderTask = null;
	}
}

// ── Plugin ──

export default class PptxViewerPlugin extends Plugin {
	async onload() {
		this.registerView(CONVERTER_VIEW, (leaf) => new PptxConverterView(leaf, this));
		this.registerView(SLIDESHOW_VIEW, () => new PdfSlideshowView(this.app.workspace.activeLeaf!));
		this.registerExtensions([PPTX_EXT], CONVERTER_VIEW);

		// When a marked PDF opens in the default viewer, switch to slideshow
		this.registerEvent(
			this.app.workspace.on("active-leaf-change", (leaf) => {
				if (leaf) this.maybeRedirectToSlideshow(leaf);
			})
		);
	}

	private async maybeRedirectToSlideshow(leaf: WorkspaceLeaf) {
		const view = leaf.view;
		if (!view || view.getViewType() !== "pdf") return;

		const file = (view as any).file as TFile | undefined;
		if (!file) return;

		const isSlideshow = await hasSlideshowMarker(file, this.app.vault);
		if (isSlideshow) {
			await leaf.setViewState({
				type: SLIDESHOW_VIEW,
				state: { file: file.path },
			});
		}
	}

	async convertAndReplace(file: TFile, leaf: WorkspaceLeaf) {
		const sofficePath = findLibreOffice();
		if (!sofficePath) {
			const frag = new DocumentFragment();
			frag.appendText("LibreOffice not found. Please ");
			const link = frag.createEl("a", { text: "install LibreOffice" });
			link.href = "https://www.libreoffice.org/download/";
			link.addEventListener("click", (e) => {
				e.preventDefault();
				window.open(link.href);
			});
			frag.appendText(" to convert PPTX files.");
			new Notice(frag, 0);
			return;
		}

		const adapter = this.app.vault.adapter as any;
		const fullPath: string = adapter.getFullPath
			? adapter.getFullPath(file.path)
			: require("path").join(adapter.basePath, file.path);

		const notice = new Notice("Converting presentation to PDF\u2026", 0);

		try {
			const pdfBuffer = await convertToPdf(fullPath, sofficePath);
			const pdfData = pdfBuffer.buffer.slice(
				pdfBuffer.byteOffset,
				pdfBuffer.byteOffset + pdfBuffer.byteLength
			);
			const pdfVaultPath = file.path.replace(/\.pptx$/i, ".pdf");
			await this.app.vault.createBinary(pdfVaultPath, pdfData);
			await this.app.vault.delete(file);

			let pdfFile = this.app.vault.getAbstractFileByPath(pdfVaultPath);
			if (!pdfFile) {
				await new Promise<void>((resolve) => {
					const ref = this.app.vault.on("create", (f) => {
						if (f.path === pdfVaultPath) {
							this.app.vault.offref(ref);
							resolve();
						}
					});
					setTimeout(() => { this.app.vault.offref(ref); resolve(); }, 2000);
				});
				pdfFile = this.app.vault.getAbstractFileByPath(pdfVaultPath);
			}

			if (pdfFile instanceof TFile) {
				// Open directly into slideshow view
				await leaf.setViewState({
					type: SLIDESHOW_VIEW,
					state: { file: pdfFile.path },
				});
			}

			notice.setMessage("Converted to PDF successfully.");
			setTimeout(() => notice.hide(), 3000);
		} catch (e) {
			notice.hide();
			new Notice(`PPTX conversion failed: ${String(e)}`, 0);
			console.error("PPTX Viewer error:", e);
		}
	}
}
