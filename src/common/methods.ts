import * as moment from "moment";

// Formats file size into human readable form
export const formatBytes = (bytes: number, decimals = 2) => {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    // Return the formatted size string
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Formats the date value
export const formatDateValue = (value: string, format: string = "MM/DD/YYYY") => {
    // Ensure a value exists
    if (value) {
        // Return the date value
        return moment(value).format(format);
    }

    // Return nothing
    return "";
}

// Formats the time value
export const formatTimeValue = (value: string, format: string = "MM/DD/YYYY HH:mm:ss") => {
    // Ensure a value exists
    if (value) {
        // Return the date value
        return moment(value).format(format);
    }

    // Return nothing
    return "";
}