#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <filesystem>
#include <algorithm>
#include <sstream>
#include <optional>
#include <cctype>

#include <xlsxwriter.h>

namespace fs = std::filesystem;

// -------- helpers --------
static std::string trim(const std::string& s){
    size_t b = s.find_first_not_of(" \t\r\n");
    size_t e = s.find_last_not_of(" \t\r\n");
    return (b==std::string::npos) ? "" : s.substr(b, e-b+1);
}
static std::string lower(std::string s){
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c){ return std::tolower(c); });
    return s;
}
static std::vector<std::string> split_tsv(const std::string& line){
    std::vector<std::string> out; std::stringstream ss(line); std::string fld;
    while(std::getline(ss, fld, '\t')) out.push_back(fld);
    return out;
}
static std::optional<double> to_number(const std::string& s){
    std::string t = trim(s); if(t.empty()) return std::nullopt;
    char* end=nullptr; double v = std::strtod(t.c_str(), &end);
    if(end && *end==0) return v; return std::nullopt;
}

// -------- data --------
struct Row {
    std::string sample;
    std::string marker;
    std::vector<std::string> alleles; // up to 8
};
using MarkerMap = std::unordered_map<std::string, Row>;        // marker -> row
using SampleMap = std::unordered_map<std::string, MarkerMap>;  // name_lower -> markers

// -------- TSV loader (tolerant) --------
static bool load_tsv(const fs::path& path, SampleMap& data, std::vector<std::string>& all_markers){
    std::ifstream in(path);
    if(!in){ std::cerr << "ERROR: cannot open " << path << "\n"; return false; }

    std::string line;
    if(!std::getline(in, line)){ std::cerr << "ERROR: empty file.\n"; return false; }

    auto hdr = split_tsv(line);
    if((int)hdr.size() < 2){
        std::cerr << "ERROR: need at least Sample name and marker columns.\n";
        return false;
    }
    int allele_cols = std::min(8, (int)hdr.size() - 2);

    std::unordered_set<std::string> marker_set;

    while(std::getline(in, line)){
        if(trim(line).empty()) continue;
        auto cols = split_tsv(line);
        if((int)cols.size() < 2) continue;
        cols.resize(2 + allele_cols, "");

        Row r;
        r.sample = trim(cols[0]);
        r.marker = trim(cols[1]);
        r.alleles.resize(8, "");
        for(int i=0;i<allele_cols;i++) r.alleles[i] = trim(cols[2+i]);

        if(r.sample.empty() || r.marker.empty()) continue;
        data[lower(r.sample)][r.marker] = r;
        marker_set.insert(r.marker);
    }

    all_markers.assign(marker_set.begin(), marker_set.end());
    std::sort(all_markers.begin(), all_markers.end());
    return true;
}

int main(){
    // ---- inputs ----
    std::string input_path_str;
    std::cout << "Enter TSV file name (with path): ";
    std::getline(std::cin, input_path_str);
    fs::path input_path = fs::path(input_path_str);

    std::string output_name;
    std::cout << "Enter output XLSX file name (e.g., result.xlsx): ";
    std::getline(std::cin, output_name);
    fs::path out_path = input_path.parent_path() / output_name;

    // ---- load once ----
    SampleMap data;
    std::vector<std::string> all_markers;
    if(!load_tsv(input_path, data, all_markers)) return 1;

    // ---- workbook ----
    lxw_workbook* wb = workbook_new(out_path.string().c_str());
    if(!wb){ std::cerr << "ERROR: cannot create xlsx.\n"; return 1; }
    lxw_worksheet* ws = workbook_add_worksheet(wb, "Comparisons");

    lxw_format* red_bg = workbook_add_format(wb);
    format_set_pattern(red_bg, LXW_PATTERN_SOLID);   // needed for visible fill
    format_set_bg_color(red_bg, LXW_COLOR_RED);

    int row = 0;
    int duo_idx = 1;

    // ---- duo loop ----
    while(true){
        std::string father, son;
        std::cout << "Enter father name (or 'end' to finish): ";
        std::getline(std::cin, father);
        if(lower(trim(father)) == "end") break;

        std::cout << "Enter son name (or 'end' to finish): ";
        std::getline(std::cin, son);
        if(lower(trim(son)) == "end") break;

        auto fit = data.find(lower(trim(father)));
        if(fit == data.end()){ std::cerr << "Father not found. Try again.\n"; continue; }
        auto sit = data.find(lower(trim(son)));
        if(sit == data.end()){ std::cerr << "Son not found. Try again.\n"; continue; }

        const auto& fmarks = fit->second;
        const auto& smarks = sit->second;

        // 1) Decide kept allele indices for THIS duo (drop columns empty for both).
        std::vector<int> keep_idx;
        for(int i=0;i<8;i++){
            bool any_val = false;
            for(const auto& m : all_markers){
                auto ifm = fmarks.find(m);
                auto ism = smarks.find(m);
                std::string f = (ifm!=fmarks.end()) ? ifm->second.alleles[i] : "";
                std::string s = (ism!=smarks.end()) ? ism->second.alleles[i] : "";
                if(!trim(f).empty() || !trim(s).empty()){ any_val = true; break; }
            }
            if(any_val) keep_idx.push_back(i);
        }
        if(keep_idx.empty()) keep_idx = {0}; // keep at least Allele 1

        // 2) Duo-specific header (alternating Father/Son per allele).
        int c0 = 0;
        worksheet_write_string(ws, row, c0++, "Duo #", nullptr);
        worksheet_write_string(ws, row, c0++, "Father", nullptr);
        worksheet_write_string(ws, row, c0++, "Son", nullptr);
        worksheet_write_string(ws, row, c0++, "Marker", nullptr);
        for(int idx : keep_idx){
            worksheet_write_string(ws, row, c0++, ("Father Allele " + std::to_string(idx+1)).c_str(), nullptr);
            worksheet_write_string(ws, row, c0++, ("Son Allele " + std::to_string(idx+1)).c_str(), nullptr);
        }
        worksheet_write_string(ws, row, c0++, "Match", nullptr);
        row++;

        // Optional width
        for(int c=0;c<c0;c++) worksheet_set_column(ws, c, c, 16.0, nullptr);

        // 3) Rows for all markers.
        for(const auto& m : all_markers){
            const Row* fr = nullptr; const Row* sr = nullptr;
            auto ifm = fmarks.find(m); if(ifm!=fmarks.end()) fr = &ifm->second;
            auto ism = smarks.find(m); if(ism!=smarks.end()) sr = &ism->second;

            int c = 0;
            worksheet_write_number(ws, row, c++, duo_idx, nullptr);
            worksheet_write_string(ws, row, c++, father.c_str(), nullptr);
            worksheet_write_string(ws, row, c++, son.c_str(), nullptr);
            worksheet_write_string(ws, row, c++, m.c_str(), nullptr);

            bool any_diff = false;
            bool stopped = false;

            // Compare kept alleles; stop at first blank on either side per marker.
            for(size_t k=0; k<keep_idx.size(); ++k){
                int i = keep_idx[k];
                std::string f = fr ? fr->alleles[i] : "";
                std::string s = sr ? sr->alleles[i] : "";

                bool stop_here = (trim(f).empty() || trim(s).empty());

                auto fn = to_number(f);
                auto sn = to_number(s);
                bool diff;
                if(fn.has_value() && sn.has_value()) diff = !(fn.value() == sn.value());
                else                                  diff = !(trim(f) == trim(s));

                // Father then Son for this allele index
                if(fn.has_value()) worksheet_write_number(ws, row, c, fn.value(), diff ? red_bg : nullptr);
                else                worksheet_write_string(ws, row, c, f.c_str(), diff ? red_bg : nullptr);
                c++;

                if(sn.has_value()) worksheet_write_number(ws, row, c, sn.value(), diff ? red_bg : nullptr);
                else                worksheet_write_string(ws, row, c, s.c_str(), diff ? red_bg : nullptr);
                c++;

                if(diff) any_diff = true;

                if(stop_here){
                    // Pad remaining kept pairs with blanks so columns align.
                    for(size_t kk=k+1; kk<keep_idx.size(); ++kk){
                        worksheet_write_string(ws, row, c++, "", nullptr); // father blank
                        worksheet_write_string(ws, row, c++, "", nullptr); // son blank
                    }
                    stopped = true;
                    break;
                }
            }

            worksheet_write_string(ws, row, c, any_diff ? "Mismatch" : "Match", nullptr);
            row++;
        }

        // spacer
        row++; duo_idx++;
    }

    if(workbook_close(wb) != LXW_NO_ERROR){
        std::cerr << "ERROR: writing xlsx.\n"; return 1;
    }
    std::cout << "Done. Saved: " << out_path << "\n";
    return 0;
}
