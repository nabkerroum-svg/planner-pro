import React, { useState, useEffect, useRef } from "react";
import { fetchAddressSuggestions } from "./nominatimSuggest.js";

const DEBOUNCE_MS = 300;
const MIN_CHARS = 3;

const defaultTheme = {
  grid: "#E8E2D8",
  pale: "#F7F2E8",
  text: "#1a1a1a",
  textMuted: "#6B6355",
  shadow: "0 4px 20px rgba(15, 10, 0, 0.12)",
};

/**
 * @param {{ value: string, onChange: (patch: { address: string, lat?: number, lon?: number }) => void, placeholder?: string, style?: object, theme?: Partial<typeof defaultTheme> }}
 */
export function AddressAutocomplete(props) {
  const theme = Object.assign({}, defaultTheme, props.theme || {});
  const [suggestions, setSuggestions] = useState([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState(null);
  const [highlight, setHighlight] = useState(-1);
  const timerRef = useRef(null);
  const wrapRef = useRef(null);

  useEffect(function () {
    function handleClick(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return function () {
      document.removeEventListener("mousedown", handleClick);
    };
  }, []);

  function scheduleFetch(text) {
    clearTimeout(timerRef.current);
    if (text.length < MIN_CHARS) {
      setSuggestions([]);
      setOpen(false);
      setLoading(false);
      setFetchError(null);
      return;
    }
    setLoading(true);
    setFetchError(null);
    timerRef.current = setTimeout(async function () {
      try {
        const list = await fetchAddressSuggestions(text);
        setSuggestions(list);
        setOpen(list.length > 0);
        setHighlight(-1);
      } catch (err) {
        setSuggestions([]);
        setOpen(false);
        setFetchError("Suggestions indisponibles");
        console.warn("[AddressAutocomplete]", err);
      } finally {
        setLoading(false);
      }
    }, DEBOUNCE_MS);
  }

  function handleChange(e) {
    const v = e.target.value;
    props.onChange({ address: v, lat: undefined, lon: undefined });
    scheduleFetch(v);
  }

  function pickItem(item) {
    props.onChange({
      address: item.displayName,
      lat: item.lat,
      lon: item.lon,
    });
    setSuggestions([]);
    setOpen(false);
    setFetchError(null);
  }

  function handleKeyDown(e) {
    if (!open || suggestions.length === 0) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setHighlight(function (h) {
        return Math.min(suggestions.length - 1, h + 1);
      });
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setHighlight(function (h) {
        return Math.max(0, h - 1);
      });
    } else if (e.key === "Enter" && highlight >= 0) {
      e.preventDefault();
      pickItem(suggestions[highlight]);
    } else if (e.key === "Escape") {
      setOpen(false);
    }
  }

  return (
    <div
      ref={wrapRef}
      style={{ position: "relative", flex: 1, minWidth: 0, ...(props.style || {}) }}
    >
      <input
        value={props.value || ""}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        onFocus={function () {
          if (suggestions.length > 0) setOpen(true);
        }}
        placeholder={props.placeholder || "Adresse..."}
        autoComplete="off"
        aria-autocomplete="list"
        aria-expanded={open}
        style={{
          padding: "8px 10px",
          borderRadius: 10,
          border: "1px solid " + theme.grid,
          width: "100%",
          boxSizing: "border-box",
          fontSize: 13,
          color: theme.text,
        }}
      />
      <div
        style={{
          position: "absolute",
          right: 10,
          top: "50%",
          transform: "translateY(-50%)",
          fontSize: 11,
          color: "#aaa",
          pointerEvents: "none",
        }}
      >
        {loading ? "..." : ""}
      </div>
      {fetchError && !loading && (
        <div style={{ fontSize: 11, color: "#c0392b", marginTop: 4 }}>{fetchError}</div>
      )}
      {open && suggestions.length > 0 && (
        <ul
          role="listbox"
          style={{
            position: "absolute",
            top: "100%",
            left: 0,
            right: 0,
            zIndex: 2000,
            margin: "4px 0 0 0",
            padding: 0,
            listStyle: "none",
            background: "#FFFFFF",
            border: "1px solid " + theme.grid,
            borderRadius: 10,
            boxShadow: theme.shadow,
            maxHeight: 240,
            overflowY: "auto",
          }}
        >
          {suggestions.map(function (item, i) {
            const sub = [item.postcode, item.city].filter(Boolean).join(" · ");
            const active = i === highlight;
            return (
              <li
                key={item.placeId + "-" + i}
                role="option"
                aria-selected={active}
                onMouseDown={function (ev) {
                  ev.preventDefault();
                  pickItem(item);
                }}
                onMouseEnter={function () {
                  setHighlight(i);
                }}
                style={{
                  padding: "10px 14px",
                  cursor: "pointer",
                  borderBottom: i < suggestions.length - 1 ? "1px solid " + theme.pale : "none",
                  background: active ? theme.pale : "#FFFFFF",
                  color: theme.text,
                }}
              >
                <div style={{ fontSize: 13, lineHeight: 1.35, fontWeight: 500 }}>{item.displayName}</div>
                {sub ? (
                  <div style={{ fontSize: 11, color: theme.textMuted, marginTop: 4 }}>{sub}</div>
                ) : null}
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
}
