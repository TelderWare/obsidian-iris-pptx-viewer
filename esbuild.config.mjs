import esbuild from "esbuild";

await esbuild.build({
	entryPoints: ["src/main.ts"],
	bundle: true,
	outfile: "main.js",
	format: "cjs",
	platform: "browser",
	external: ["obsidian", "fs", "fs/promises", "path", "os", "child_process"],
	target: "es2020",
	sourcemap: false,
	minify: false,
	logLevel: "info",
});
