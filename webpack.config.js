/* eslint-disable no-undef */

require("dotenv").config(); // loads .env into process.env for the DefinePlugin injection below

const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const PROD_BASE = "https://goindia-addin.netlify.app/"; // must match manifest.xml's own hardcoded SourceLocation domain

// Single source of truth for the version: manifest.xml's own <Version> tag (the same value you
// already bump by hand for every Partner Center submission) — not a separate package.json field
// that could drift out of sync with it.
function getManifestVersion() {
  const manifestContent = fs.readFileSync(path.resolve(__dirname, "manifest.xml"), "utf8");
  const match = manifestContent.match(/<Version>([\d.]+)<\/Version>/);
  if (!match) throw new Error("Could not find <Version> in manifest.xml");
  return match[1];
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

// CopyWebpackPlugin with a `to` that escapes output.path (needed here since legacy-root must land
// at the true dist/ root while everything else builds into a versioned subfolder) hits a real,
// documented edge case — its asset tracking assumes everything lives under output.path, and an
// absolute `to` outside it can silently collide with webpack's own asset registry instead of doing
// a clean copy (confirmed empirically: the file that landed at dist/taskpane.js was neither
// legacy-root's content nor this build's own compiled output). Plain fs.cpSync after emit has no
// knowledge of webpack's asset tracking at all, so it can't collide with it.
class RestoreLegacyRootPlugin {
  constructor({ from, to }) {
    this.from = from;
    this.to = to;
  }
  apply(compiler) {
    compiler.hooks.afterEmit.tap("RestoreLegacyRootPlugin", () => {
      if (!fs.existsSync(this.from)) return;
      fs.cpSync(this.from, this.to, { recursive: true });
    });
  }
}

// Copies specific build outputs from this version's folder up to the TRUE dist/ root. Needed for
// the Google sign-in pages: google-signin-dialog.html builds its redirect_uri as
// `${location.origin}/google-signin-callback.html` — root-absolute, not relative — and that exact
// unversioned path is what is registered in Google Cloud as an Authorized redirect URI. If those
// two files only ever landed in dist/v{VERSION}/, the OAuth redirect would 404 on every release,
// and pointing them at a versioned path instead would mean re-registering a new redirect URI in
// Google Cloud for every single release.
//
// Deliberately NOT a CopyWebpackPlugin pattern with an absolute `to`: that escapes output.path and
// hits the asset-registry collision documented on RestoreLegacyRootPlugin above. Same fs.cpSync
// after emit approach, for the same reason.
//
// Must be registered AFTER RestoreLegacyRootPlugin — that one recursively copies the frozen
// snapshot over dist/ root, and these files have to survive it. (It wouldn't actually clobber them
// today, since legacy-root contains no google-signin*.html, but ordering shouldn't depend on that.)
// No-op when output.path already IS distRoot, i.e. every dev build.
class CopyToDistRootPlugin {
  constructor({ files, from, to }) {
    this.files = files;
    this.from = from;
    this.to = to;
  }
  apply(compiler) {
    compiler.hooks.afterEmit.tap("CopyToDistRootPlugin", () => {
      if (path.resolve(this.from) === path.resolve(this.to)) return;
      fs.mkdirSync(this.to, { recursive: true });
      for (const file of this.files) {
        const src = path.join(this.from, file);
        if (!fs.existsSync(src)) continue;
        fs.cpSync(src, path.join(this.to, file));
      }
    });
  }
}

// Both legacy-root/taskpane.js AND every finalized dist/v*/ release folder are git-tracked (so they
// survive every future clean CI checkout), which means none of them can ever contain the real
// OPENROUTER_KEY committed in plain text — that would put the live production key into git history.
// Each committed copy holds a placeholder instead; this plugin re-injects the real value (from
// .env, same source DefinePlugin already uses for the CURRENT build's own key) into every
// taskpane.js under dist/ after emit — dist/taskpane.js (the restored root) plus every
// dist/v*/taskpane.js inherited as-is from a past commit (any version other than the one just
// freshly compiled, which never contains the placeholder to begin with, making this a safe no-op
// for it). Runs on EVERY build, so it keeps working for whatever version gets committed next.
class ReinjectSecretsPlugin {
  constructor({ distRoot, secretPlaceholder, secretValue }) {
    this.distRoot = distRoot;
    this.secretPlaceholder = secretPlaceholder;
    this.secretValue = secretValue;
  }
  apply(compiler) {
    compiler.hooks.afterEmit.tap("ReinjectSecretsPlugin", () => {
      if (!this.secretPlaceholder || !fs.existsSync(this.distRoot)) return;
      const candidates = [path.join(this.distRoot, "taskpane.js")];
      for (const entry of fs.readdirSync(this.distRoot, { withFileTypes: true })) {
        if (entry.isDirectory() && entry.name.startsWith("v")) {
          candidates.push(path.join(this.distRoot, entry.name, "taskpane.js"));
        }
      }
      for (const file of candidates) {
        if (!fs.existsSync(file)) continue;
        const content = fs.readFileSync(file, "utf8");
        if (content.includes(this.secretPlaceholder)) {
          fs.writeFileSync(file, content.split(this.secretPlaceholder).join(this.secretValue || ""));
        }
      }
    });
  }
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const VERSION = getManifestVersion();

  // Dev builds stay exactly as before — unversioned, served from dist/ root at localhost, no
  // change to the local sideload/test workflow. Only production builds go versioned, since only
  // production builds are what a stale, already-approved Partner Center manifest could ever hit.
  const outputPath = dev ? path.resolve(__dirname, "dist") : path.resolve(__dirname, "dist", "v" + VERSION);
  const distRoot = path.resolve(__dirname, "dist");

  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      path: outputPath,
      clean: true, // only ever cleans outputPath itself (this version's own folder) — never touches
                   // sibling dist/v*/ folders from earlier releases, or the frozen dist/ root below.
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      // Injects secrets from .env (never committed) as build-time string literals. This keeps the
      // raw key out of the git-tracked source file — taskpane.js references
      // process.env.OPENROUTER_KEY, and webpack replaces that exact expression with the real value
      // at build time. NOTE: the built dist/ output still contains the real value baked in (there's
      // no server here to keep it secret from the browser at runtime) — dist/ must stay
      // gitignored and treated as sensitive; this only fixes what gets committed to source control.
      new webpack.DefinePlugin({
        "process.env.OPENROUTER_KEY": JSON.stringify(process.env.OPENROUTER_KEY || ""),
        // Only ever used for the ut@gmail.com certification test account (UserId "test123"),
        // which isn't a real DB user so /dbcatalog/get-mcp-key can't resolve it — see
        // ensureMcpAccess() in taskpane.js. Same caveat as OPENROUTER_KEY: this still ends up
        // baked into the shipped dist/taskpane.js bundle, there's no way around that client-side.
        "process.env.TEST_MCP_KEY": JSON.stringify(process.env.TEST_MCP_KEY || ""),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "EULA.html",
            to: "[name][ext]",
          },
          {
            from: "privacy.html",
            to: "[name][ext]",
          },
          {
            from: "index.html",
            to: "[name][ext]",
          },
          {
            // Google sign-in dialog pages (Office.context.ui.displayDialogAsync from
            // taskpane.html's startGoogleSignIn). This `to` is relative to output.path, so on a
            // production build these land in dist/v{VERSION}/ along with everything else —
            // CopyToDistRootPlugin below is what additionally lifts them to the unversioned dist/
            // root, which is where Google's registered Authorized redirect URI actually points.
            from: "google-signin-dialog.html",
            to: "[name][ext]",
          },
          {
            from: "google-signin-callback.html",
            to: "[name][ext]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) return content;
              // Points this version's OWN manifest copy at its own versioned folder — this is the
              // file you'd actually submit to Partner Center for THIS release. Self-contained on
              // purpose: also re-points icon URLs into this version's folder, trading a little
              // duplicated icon weight per release for zero fragile selective-URL matching.
              return content.toString().split(PROD_BASE).join(`${PROD_BASE}v${VERSION}/`);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      // Restores the frozen, currently-live production snapshot to the TRUE dist/ root (not this
      // version's own versioned subfolder) after every build — this is what keeps whatever manifest
      // Partner Center already has approved working, no matter what changes in src/ going forward.
      // Then re-injects the real OPENROUTER_KEY into every placeholder-holding taskpane.js under
      // dist/ (root + any inherited past-release folders) — see ReinjectSecretsPlugin above.
      // Skipped in dev: irrelevant to local sideload testing.
      ...(dev
        ? []
        : [
            new RestoreLegacyRootPlugin({
              from: path.resolve(__dirname, "legacy-root"),
              to: distRoot,
            }),
            new CopyToDistRootPlugin({
              files: ["google-signin-dialog.html", "google-signin-callback.html"],
              from: outputPath,
              to: distRoot,
            }),
            new ReinjectSecretsPlugin({
              distRoot,
              secretPlaceholder: "__REDACTED_OPENROUTER_KEY__",
              secretValue: process.env.OPENROUTER_KEY || "",
            }),
          ]),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
