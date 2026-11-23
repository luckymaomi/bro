(function () {
    "use strict";

    var storageKey = "bro.theme";
    // 旧仓主题键：皮燕子 → 宝宝 → 小弟（bro）；读时兼容，写回只落现行键
    var legacyStorageKeys = ["baobao.theme", "piyanzi.theme"];
    var defaultTheme = document.documentElement.dataset.defaultTheme === "dark" ? "dark" : "light";
    function normalizeTheme(value) {
        return value === "dark" || value === "light" ? value : defaultTheme;
    }

    function readStoredTheme() {
        try {
            var raw = window.localStorage.getItem(storageKey);
            if (raw === "dark" || raw === "light") return raw;
            for (var i = 0; i < legacyStorageKeys.length; i += 1) {
                var legacy = window.localStorage.getItem(legacyStorageKeys[i]);
                if (legacy === "dark" || legacy === "light") {
                    window.localStorage.setItem(storageKey, legacy);
                    return legacy;
                }
            }
            return defaultTheme;
        } catch (error) {
            return defaultTheme;
        }
    }

    function storeTheme(theme) {
        try {
            window.localStorage.setItem(storageKey, theme);
            for (var i = 0; i < legacyStorageKeys.length; i += 1) {
                window.localStorage.removeItem(legacyStorageKeys[i]);
            }
        } catch (error) {
            return;
        }
    }

    function applyTheme(theme, shouldStore) {
        var nextTheme = normalizeTheme(theme);
        var bootstrapTheme = nextTheme;
        var root = document.documentElement;

        root.dataset.theme = nextTheme;
        root.dataset.bsTheme = bootstrapTheme;
        root.style.colorScheme = bootstrapTheme;

        if (document.body) {
            document.body.dataset.theme = nextTheme;
        }

        if (shouldStore) {
            storeTheme(nextTheme);
        }

        window.dispatchEvent(new CustomEvent("bro:themechange", {
            detail: { theme: nextTheme }
        }));

        return nextTheme;
    }

    function getTheme() {
        return normalizeTheme(document.documentElement.dataset.theme || readStoredTheme());
    }

    function setTheme(theme) {
        return applyTheme(theme, true);
    }

    function toggleTheme() {
        return setTheme(getTheme() === "dark" ? "light" : "dark");
    }

    window.BroTheme = {
        getTheme: getTheme,
        setTheme: setTheme,
        toggleTheme: toggleTheme,
        storageKey: storageKey
    };

    applyTheme(readStoredTheme(), false);

    document.addEventListener("DOMContentLoaded", function () {
        applyTheme(getTheme(), false);
    });
}());
