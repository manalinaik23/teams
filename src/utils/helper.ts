export const padZeroesLeft = (num, size) => {
        var s = num + "";
        while (s.length < size) s = "0" + s;
        return s;
};

export const SortArrayByPeriod = (order) => {
        var sort_order = 1;
        if (order === "desc") {
                sort_order = -1;
        }
        return (a, b) => {
                // a should come before b in the sorted order
                if (new Date(a["Year"], a["Month"], 1) < new Date(b["Year"], b["Month"], 1)) {
                        return -1 * sort_order;
                        // a should come after b in the sorted order
                } else if (new Date(a["Year"], a["Month"], 1) > new Date(b["Year"], b["Month"], 1)) {
                        return 1 * sort_order;
                        // a and b are the same
                } else {
                        return 0 * sort_order;
                }
        };
};

export const includes=(pattern, text)=> {
        if (pattern.length == 0)
                return 0; // Immediate match
        pattern= pattern.toLowerCase();
        text = text ? text.toLowerCase() :"";
        // Compute longest suffix-prefix table
        var lsp = [0]; // Base case
        for (var i = 1; i < pattern.length; i++) {
                var j = lsp[i - 1]; // Start by assuming we're extending the previous LSP
                while (j > 0 && pattern.charAt(i) != pattern.charAt(j))
                        j = lsp[j - 1];
                if (pattern.charAt(i) == pattern.charAt(j))
                        j++;
                lsp.push(j);
        }

        // Walk through text string
        var k = 0; // Number of chars matched in pattern
        for (var l = 0; l < text.length; l++) {
                while (k > 0 && text.charAt(l) != pattern.charAt(k))
                        k = lsp[k - 1]; // Fall back in the pattern
                if (text.charAt(l) == pattern.charAt(k)) {
                        k++; // Next char matched, increment position
                        if (k == pattern.length)
                                return l - (k - 1);
                }
        }
        return -1; // Not found
};