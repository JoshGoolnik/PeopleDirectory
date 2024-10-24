import * as React from "react";

export function AvailableIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#92c353">
                <circle cx="5" cy="5" r="5"/>
                <polyline points="2,4 4,7 8,3" fill = "none" stroke = "white" stroke-width = "1"/>
        </svg>
    );
}

export function AwayIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#fcd116">
            <circle cx="5" cy="5" r="5"/>
            <polyline points="5,3 5,6 7,7 " fill = "none" stroke = "white" stroke-width = "1"/>
        </svg>
    );
}

export function BusyIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#c4314b">
            <circle cx="5" cy="5" r="5"/>
        </svg>
    );
}

export function DoNotDisturbIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" fill="#c4314b">
            <circle cx="5" cy="5" r="5"/>
            <line x1="1" y1="5" x2="9" y2="5" stroke = "white" stroke-width = "2" />
        </svg>
    );
}

export function OutOfOfficeIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" stroke="#b4009e" fill="none">
            <circle cx="5" cy="5" r="5"/>
            <polyline points="5,2 2,5 5,8"/>
            <line x1 = "2" y1="5" x2="8"  y2="5"/>
        </svg>
    );
}

export function OfflineIcon() {
    return (
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="13" viewBox="0 0 10 10" stroke="#959595" fill="none">
            <circle cx="5" cy="5" r="5"/>
            <line x1 = "2" y1="2" x2="8"  y2="8" />
            <line x1 = "2" y1="8" x2="8"  y2="2" />
        </svg>
    );
}