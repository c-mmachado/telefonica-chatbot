// Imports util to set the default depth for object inspection when logging
import * as util from "util";
util.inspect.defaultOptions.depth = null;

import { ILogObj, IMeta, ISettings, IStackFrame, Logger } from "tslog";
// import { prettyLogStyles } from "tslog/dist/types/prettyLogStyles";
// import { formatTemplate } from "tslog/dist/types/formatTemplate";
import chalk from "chalk";

import { maskObject } from "./misc";
import { types } from "util";
import { normalize } from "path";

/**
 * Stores the original console before being overridden by the logger so that it can still be used within the logger transport.
 *
 * @private
 */
const originalConsole = { ...console };

/**
 * The root logger for the application.
 *
 * This logger is configured to provide pretty-printed logs with masking of sensitive information, such as passwords and secrets.
 * It overrides the default console methods to ensure all logs go through this logger for consistent formatting and masking.
 *
 * @public
 */
export const logger: Logger<ILogObj> = new Logger({
    type: "pretty",
    name: "RootLogger",
    hideLogPositionForProduction: false,
    prettyLogTemplate: "[{{dateIsoStr}}] [{{fileNameWithLine}}] [{{name}}] [{{logLevelName}}] ",
    prettyErrorTemplate: "{{errorName}}: {{errorMessage}}\n{{errorStack}}",
    prettyErrorStackTemplate: "  • at {{method}} ({{filePathWithLine}})",
    prettyInspectOptions: {
        depth: null,
    },
    prettyLogStyles: {
        logLevelName: {
            "*": ["bold", "black", "bgWhite"],
            SILLY: ["bold", "black"],
            TRACE: ["bold", "black"],
            DEBUG: ["bold", "cyan"],
            INFO: ["bold", "blue"],
            WARN: ["bold", "yellow"],
            ERROR: ["bold", "red"],
            FATAL: ["bold", "white", "bgRed"],
        },
        dateIsoStr: ["dim", "magenta"],
        name: ["dim", "black"],
        nameWithDelimiterPrefix: ["dim", "cyan"],
        nameWithDelimiterSuffix: ["dim", "cyan"],
        errorName: ["bold", "white", "bgRed"],
        filePathWithLine: ["bold", "black"],
        fileNameWithLine: ["yellow"],
        fileName: ["yellow"],
        fileLine: ["yellow"],
        filePath: ["yellow"],
        location: ["yellow"],
        runtime: ["dim", "green"],
        runtimeVersion: ["dim", "green"],
    } as any,
    overwrite: {
        mask: mask,
        formatLogObj: formatLogObj,
        addPlaceholders: (
            logObjMeta: IMeta & Partial<{ runtimeVersion: string }>,
            placeholderValues: Record<string, string | number>
        ): void => {
            // console.log("addPlaceholders called with logObjMeta", logObjMeta);
            // console.log("placeholderValues", placeholderValues);

            const location: string = `${placeholderValues.filePathWithLine}`; // @${logObjMeta.path?.method || "<unknown>"}
            // const locationLength = placeholderValues.location.length;
            // let leftPadding = (50 - locationLength) / 2;
            // let rightPadding = leftPadding;
            // if (locationLength % 2 == 0) {
            //   rightPadding = leftPadding - 1;
            // }
            // leftPadding = leftPadding < 0 ? 0 : leftPadding;
            // rightPadding = rightPadding < 0 ? 0 : rightPadding;
            // placeholderValues.location = `${placeholderValues.location.padStart(
            //   leftPadding + locationLength < 50 ? leftPadding + locationLength : 50
            // )}${" ".repeat(rightPadding)}`;
            placeholderValues.fileNameWithLine = `${location.padStart(50)}`;

            const levelLength = logObjMeta.logLevelName.length;
            let leftPadding = (8 - levelLength) / 2;
            let rightPadding = leftPadding;
            if (levelLength % 2 == 0) {
                rightPadding = leftPadding - 1;
            }
            leftPadding = leftPadding < 0 ? 0 : leftPadding;
            rightPadding = rightPadding < 0 ? 0 : rightPadding;
            placeholderValues.logLevelName = `${logObjMeta.logLevelName.padStart(
                leftPadding + levelLength < 8 ? leftPadding + levelLength : 8
            )}${" ".repeat(rightPadding)}`;
        },
        transportFormatted: (
            logMetaMarkup: string,
            logArgs: unknown[],
            logErrors: string[],
            _logMeta: IMeta | undefined,
            settings: ISettings<ILogObj> | undefined
        ): void => {
            const logErrorsStr = (logErrors.length > 0 && logArgs.length > 0 ? "\n" : "") + logErrors.join("\n");
            if (!settings) {
                originalConsole.log(logMetaMarkup + util.format(...logArgs) + logErrorsStr);
                return;
            }

            settings.prettyInspectOptions.colors = settings?.stylePrettyLogs;
            originalConsole.log(
                logMetaMarkup + util.formatWithOptions(settings?.prettyInspectOptions ?? {}, ...logArgs) + logErrorsStr
            );
        },
    },
});
// (logger as any).runtime.prettyFormatLogObj = prettyFormatLogObj;

/**
 * Masks sensitive information in the provided arguments.
 *
 * Function responsible for recursively traversing `args` and making keys that match a regular expression,
 * replacing their values with asterisks (`******`).
 *
 * @param args An array of arguments to be masked.
 * @returns The masked array of arguments.
 *
 * @private
 */
function mask(args: unknown[]): unknown[] {
    for (let i = 0; i < args.length; i++) {
        if (args[i] === null || args[i] === undefined || args[i] instanceof Error) {
            // Skip null or undefined values and Error instances
            continue;
        } else if (typeof args[i] === "object" || args[i] instanceof Object) {
            args[i] = maskObject(args[i], /.*(?:password|secret|token|authorization).*/gi, "******");
        } else if (Array.isArray(args[i])) {
            args[i] = mask(args[i] as unknown[]);
        }
    }
    return (logger as any)._mask(args);
}

/**
 * Formats the log object by separating errors from other arguments and applying color formatting.
 *
 * @param maskedArgs An array of masked arguments to be formatted.
 * @param settings The logger settings used for formatting.
 * @returns An object containing the formatted `args` and `errors`.
 *
 * @private
 */
function formatLogObj(maskedArgs: unknown[], settings: ISettings<ILogObj>): { args: unknown[]; errors: string[] } {
    // console.log("formatLogObj called with maskedArgs", maskedArgs);
    // console.log("settings", settings);
    // for (let i = 0; i < maskedArgs.length; i++) {
    //   let arg: unknown = maskedArgs[i];
    //   maskedArgs[i] = _colorString(arg);
    // }
    // TODO: This call will format error messages by merging all its properties into a single string, modify it to only show error name, error message and stack trace
    // return (logger as any).runtime.prettyFormatLogObj(maskedArgs, settings);

    return maskedArgs.reduce(
        (result: { args: unknown[]; errors: string[] }, arg: unknown) => {
            isError(arg)
                ? result.errors.push(prettyFormatErrorObj(arg, settings))
                : result.args.push(_colorString(arg));
            return result;
        },
        { args: [], errors: [] }
    );
}

/**
 * Checks if the provided value is an {@link Error} instance.
 *
 * @param e The value to check.
 * @returns `true` if the value is an {@link Error} instance, `false` otherwise.
 */
function isError(e: unknown): e is Error {
    return types?.isNativeError != null ? types.isNativeError(e) : e instanceof Error;
}

// function prettyFormatLogObj(
//   maskedArgs: unknown[],
//   settings: ISettings<ILogObj>
// ): { args: unknown[]; errors: string[] } {
//   return maskedArgs.reduce(
//     (result: { args: unknown[]; errors: string[] }, arg: unknown) => {
//       isError(arg)
//         ? result.errors.push(prettyFormatErrorObj(arg, settings))
//         : result.args.push(arg);
//       return result;
//     },
//     { args: [], errors: [] }
//   );
// }

/**
 * Formats an {@link Error} object into a pretty string representation.
 *
 * @param error The {@link Error} object to format.
 * @param settings The logger settings used for formatting.
 * @returns A formatted string representation of `error`.
 *
 * @private
 */
function prettyFormatErrorObj(error: Error, settings: ISettings<ILogObj>): string {
    const errorStackStr = getErrorTrace(error).map((stackFrame: IStackFrame) => {
        return _formatTemplate(settings, settings.prettyErrorStackTemplate, { ...stackFrame }, true);
    });
    const placeholderValuesError = {
        errorName: ` ${error.name} `,
        errorMessage: _colorString(error.message),
        // Object.getOwnPropertyNames(error)
        //   .reduce((result: string[], key: string) => {
        //     if (key !== "stack") {
        //       result.push(error[key as keyof Error] as string);
        //     }
        //     return result;
        //   }, [])
        //   .join(", "),
        errorStack: errorStackStr.join("\n"),
    };
    return _formatTemplate(settings, settings.prettyErrorTemplate, placeholderValuesError);
}

/**
 * Extracts the stack trace from an {@link Error} object.
 *
 * @param error The {@link Error} object from which to extract the stack trace.
 * @returns An array of {@link IStackFrame} representing the stack trace.
 *
 * @private
 */
function getErrorTrace(error: Error): IStackFrame[] {
    const stackFrames = error?.stack?.split("\n")?.reduce((result: IStackFrame[], line: string): IStackFrame[] => {
        if (line.includes("    at ")) {
            result.push(stackLineToStackFrame(line));
        }
        return result;
    }, []);
    return stackFrames!;
}

/**
 * Parses a single line of an error stack trace into an {@link IStackFrame} object.
 *
 * @param line The line of the stack trace to parse.
 * @returns An {@link IStackFrame} object representing the parsed stack frame.
 *
 * @private
 */
function stackLineToStackFrame(line: string): IStackFrame {
    const pathResult: IStackFrame = {
        fullFilePath: undefined,
        fileName: undefined,
        fileNameWithLine: undefined,
        fileColumn: undefined,
        fileLine: undefined,
        filePath: undefined,
        filePathWithLine: undefined,
        method: undefined,
    };
    if (line != null && line.includes("    at ")) {
        line = line.replace(/^\s+at\s+/gm, "");
        const errorStackLine = line.split(" (");
        const fullFilePath = line?.slice(-1) === ")" ? line?.match(/\(([^)]+)\)/)?.[1] : line;
        const pathArray = fullFilePath?.includes(":")
            ? fullFilePath?.replace("file://", "")?.replace(process.cwd(), "")?.split(":")
            : undefined;
        const fileColumn = pathArray?.pop();
        const fileLine = pathArray?.pop();
        const filePath = pathArray?.pop();
        const filePathWithLine = normalize(`${filePath}:${fileLine}`);
        const fileName = filePath?.split("/")?.pop();
        const fileNameWithLine = `${fileName}:${fileLine}`;
        if (filePath != null && filePath.length > 0) {
            pathResult.fullFilePath = fullFilePath;
            pathResult.fileName = fileName;
            pathResult.fileNameWithLine = fileNameWithLine;
            pathResult.fileColumn = fileColumn;
            pathResult.fileLine = fileLine;
            pathResult.filePath = filePath;
            pathResult.filePathWithLine = filePathWithLine;
            pathResult.method = errorStackLine?.[1] != null ? errorStackLine?.[0] : undefined;
        }
    }
    return pathResult;
}

/**
 * Type representing styles for log formatting.
 *
 * @private
 */
type Style = string | string[] | Record<string, string | string[]>;

/**
 * Formats a template string by replacing placeholders with corresponding values and applying styles.
 *
 * Supports ANSI color codes for styling when `settings.stylePrettyLogs` is enabled in the logger settings and
 * applies styles defined in `settings.prettyLogStyles`.
 *
 * Placeholders in the template are denoted by `{{placeholderName}}`
 * and are replaced with values from the `values` object. If a placeholder does not have a corresponding value and
 * `hideUnsetPlaceholder` is true, it is replaced with an empty string; otherwise, it remains unchanged.
 *
 * Example:
 * ```ts
 * const settings: ISettings<ILogObj> = {
 *   stylePrettyLogs: true,
 *  prettyLogStyles: {
 *    name: "green",
 *  },
 * };
 *
 * const template = "Hello, {{name}}!";
 * const values = { name: "World" };
 * const result = _formatTemplate(settings, template, values);
 * // result: "Hello, \u001b[32mWorld\u001b[39m!"
 * ```
 *
 * @param settings The logger settings used for formatting.
 * @param template The template string containing placeholders.
 * @param values The values to replace placeholders with.
 * @param hideUnsetPlaceholder Whether to hide placeholders that do not have corresponding values.
 * @returns The formatted string with placeholders replaced and styles applied.
 *
 * @private
 */
function _formatTemplate(
    settings: ISettings<ILogObj>,
    template: string | undefined,
    values: Record<string, unknown>,
    hideUnsetPlaceholder: boolean = false
) {
    const templateString = String(template);

    const ansiColorWrap = (placeholderValue: string, code: number[]) =>
        `\u001b[${code[0]}m${placeholderValue}\u001b[${code[1]}m`;

    const styleWrap = (value: string, style: Style): string => {
        if (style != null && typeof style === "string") {
            return ansiColorWrap(value, (prettyLogStyles as any)[style]);
        } else if (style != null && Array.isArray(style)) {
            return style.reduce((prevValue: string, thisStyle: string) => styleWrap(prevValue, thisStyle), value);
        } else {
            if (style != null && style[value.trim()] != null) {
                return styleWrap(value, style[value.trim()]);
            } else if (style != null && style["*"] != null) {
                return styleWrap(value, style["*"]);
            } else {
                return value;
            }
        }
    };

    const defaultStyle = null;
    return templateString.replace(/{{(.+?)}}/g, (_: string, placeholder: string) => {
        const value = values[placeholder] != null ? String(values[placeholder]) : hideUnsetPlaceholder ? "" : _;
        return settings.stylePrettyLogs
            ? styleWrap(value, (settings?.prettyLogStyles as any)?.[placeholder] ?? defaultStyle) +
                  ansiColorWrap("", prettyLogStyles.reset)
            : value;
    });
}

/**
 * Applies color formatting to strings, URLs, file paths, and quoted strings within the provided argument.
 *
 * @param arg The argument to apply color formatting to.
 * @returns The color-formatted string or the original argument if no formatting is applied.
 *
 * @private
 */
function _colorString(arg: unknown): string | unknown {
    if (arg instanceof Error) {
        arg.message = _colorString(arg.message) as string;
    } else if (typeof arg === "string") {
        const johnGruberURLRegEx =
            /\b((?:[a-z][\w-]+:(?:\/{1,3}|[a-z0-9%])|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}\/)(?:[^\s()<>]+|(?:\([^\s()<>]+|(?:\([^\s()<>]+\)))*\))+(?:(?:\([^\s()<>]+|(?:\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))/;
        const filePathRegEx =
            /((?:\b[a-zA-Z]:[\/\\]*|[~\.]?[\/\\]+)[\w.-]+(?:[\/\\]+[\w.-]+)*|[\w.-]+(?:[\/\\]+[\w.-]+)+)/;
        const quotedStringRegEx = /('[^']*')/;

        // Highlight URLs in blue and file paths in yellow in a single pass to avoid overlapping replacements
        arg = arg.replace(
            new RegExp(`${johnGruberURLRegEx.source}|${filePathRegEx.source}|${quotedStringRegEx.source}`, "gi"),
            (
                _match: string,
                url: string | undefined,
                path: string | undefined,
                quotedStr: string | undefined,
                _offset: number,
                _fullStr: string
            ) => {
                return quotedStr ? chalk.green(quotedStr) : url ? chalk.blue(url) : chalk.yellow(path);
            }
        );
        return arg;
    } else if (typeof arg === "symbol") {
        return chalk`{blue ${arg}}`;
    }
    return arg;
}

/**
 * Available keyword based styles for pretty log formatting.
 *
 * @private
 */
const prettyLogStyles = {
    reset: [0, 0],
    bold: [1, 22],
    dim: [2, 22],
    italic: [3, 23],
    underline: [4, 24],
    overline: [53, 55],
    inverse: [7, 27],
    hidden: [8, 28],
    strikethrough: [9, 29],
    black: [30, 39],
    red: [31, 39],
    green: [32, 39],
    yellow: [33, 39],
    blue: [34, 39],
    magenta: [35, 39],
    cyan: [36, 39],
    white: [37, 39],
    blackBright: [90, 39],
    redBright: [91, 39],
    greenBright: [92, 39],
    yellowBright: [93, 39],
    blueBright: [94, 39],
    magentaBright: [95, 39],
    cyanBright: [96, 39],
    whiteBright: [97, 39],
    bgBlack: [40, 49],
    bgRed: [41, 49],
    bgGreen: [42, 49],
    bgYellow: [43, 49],
    bgBlue: [44, 49],
    bgMagenta: [45, 49],
    bgCyan: [46, 49],
    bgWhite: [47, 49],
    bgBlackBright: [100, 49],
    bgRedBright: [101, 49],
    bgGreenBright: [102, 49],
    bgYellowBright: [103, 49],
    bgBlueBright: [104, 49],
    bgMagentaBright: [105, 49],
    bgCyanBright: [106, 49],
    bgWhiteBright: [107, 49],
};

// Override console methods
console.log = logger.info.bind(logger);
console.info = logger.info.bind(logger);
console.warn = logger.warn.bind(logger);
console.error = logger.error.bind(logger);
console.debug = logger.debug.bind(logger);
console.trace = logger.trace.bind(logger);

// logger.debug(
//   42,
//   42.0,
//   "'Helloworld'",
//   "Foo Bar",
//   true,
//   chalk`{red ${typeof "true"}}`,
//   Object,
//   [1, 2, 3],
//   { key: "value", anotherKey: 123 }
// );
// logger.debug("This is a debug message");

// logger.warn("This is a warning message");

// logger.info("This is an info message");

// logger.error("This is an error message");
// logger.error(new Error("This is a 'test' error"));

// logger.fatal("Fatal error occurred");
// logger.fatal(new ErrorWithCode("This is a fatal error"));
// logger.fatal(
//   "This is a fatal error with additional fields",
//   new ErrorWithCode("Fatal error details")
// );
