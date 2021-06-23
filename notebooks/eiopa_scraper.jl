### A Pluto.jl notebook ###
# v0.14.8

using Markdown
using InteractiveUtils

# This Pluto notebook uses @bind for interactivity. When running this notebook outside of Pluto, the following 'mock version' of @bind gives bound variables a default value (instead of an error).
macro bind(def, element)
    quote
        local el = $(esc(element))
        global $(esc(def)) = Core.applicable(Base.get, el) ? Base.get(el) : missing
        el
    end
end

# ╔═╡ 282cd541-198e-4419-926d-13bb25826fc3
md"""
# Scraping risk-free interest rate term structures
EIOPA, the European Insurance and Occupational Pensions Authority, 
is a European Union financial regulatory institution. In this capacity 
it handles insurance and occupational pensions supervision in the EU. 
Besides supervision it also publishes tools and data which can be used 
by - among others - insurance companies. An example of data which is 
published (monthly) are the risk-free interest rate (RFR) term structures. 
The purpose of these term structures is to ensure a consistent calculation 
of technical provisions by companies across Europe.
"""

# ╔═╡ a5bc5425-0972-47cc-be22-9aeabf2345bb
md"""
In this notebook the pages on the EIOPA website which provide the risk-free 
interest rate term structures (current and previous versions) are scraped: 
zip-files containing Excel files with the risk-free interest rate term structures 
are identified by their name (date at the end of a month, *yyyymmdd*) and 
downloaded: from a downloaded zip-file the risk-free rate term structures is read 
from one of the Excel files. Selected term structures can be shown in a plot.
"""

# ╔═╡ 849e4709-1dea-4cae-a7ff-34ff43f6a613
html"""
The following pages are relevant on EIOPA's website:</br>
<ul>
<li>
<a href="https://www.eiopa.europa.eu/tools-and-data/risk-free-interest-rate-term-structures_en" target="_blank" rel="noopener noreferrer">Current EIOPA risk-free interest rate term structures</a>.
</li>
<li>
<a href="https://www.eiopa.europa.eu/risk-free-rate-previous-releases-and-preparatory-phase" target="_blank" rel="noopener noreferrer">Previous EIOPA risk-free interest rate term structures</a>.
</li>
</ul>
These are the pages that will be scraped.
"""

# ╔═╡ 9e7d2f3d-c45e-427b-bbde-f3d00f55a3b4
md"""
# Plot risk-free interest rate term structures
"""

# ╔═╡ 7c527a88-7eb7-4c6c-b14c-c6f22ea0083e
md"""
# Imports
"""

# ╔═╡ 0d6afd42-c7ca-11eb-27d6-1db1a07a7e73
begin
    import HTTP
    import Gumbo
    import Cascadia
    import Downloads
    import ZipFile
    import XLSX
    import Plots
    import PlutoUI
end

# ╔═╡ e77d8d03-c9b0-4ed6-ba64-fd7ab03f98db
md"""
Files with RFR term structures can be downloaded one by one as 
they are selected to be plotted, or they can be downloaded all 
at once using the checkbox below. Once files are downloaded, this 
local cache will be used when a specific RFR term structure is 
selected (although it is possible to download a file again and 
overwrite the local file).\
\
Download all zip-files with RFR term structures?\
$(@bind download_all PlutoUI.CheckBox(default=false))\
*(Note: this may take several seconds).*
"""

# ╔═╡ f57e9bfc-7b64-4460-ac3c-05573a32859c
md"""
Select the number of years for which the interest rates should be shown. 
The minimum is 30 years, the maximum is 150 years:\
$(@bind rfr_upper_limit PlutoUI.Slider(30:1:150;default=60, show_value=true))
"""

# ╔═╡ 1322974d-16bd-4fea-853a-335c77d1721a
md"""
# Main program
"""

# ╔═╡ efc29f37-e649-4d5c-beaf-082c7c11f9f3
md"""
This section contains the code for scraping the EIOPA website and 
processing the data. Supporting functions that are used can be found 
in the next section.

First specify the pages with risk-free interest rate structures (current / previous):
"""

# ╔═╡ d8399f9d-8984-41d6-970a-cf7d65b2cac1
begin
    rfr_urls = Vector{String}()
    # Recent rfr's
    push!(rfr_urls, "https://www.eiopa.europa.eu/tools-and-data/risk-free-interest-rate-term-structures_en")
    # Previous rfr's
    push!(rfr_urls, "https://www.eiopa.europa.eu/risk-free-rate-previous-releases-and-preparatory-phase")
end;

# ╔═╡ 8fcc98e7-1036-4cbd-8ce1-d745d2c50484
md"""
Specify the directories where downloaded zip-files and unzipped Excel 
files with RFR term structures will be saved:
"""

# ╔═╡ d8fbcacf-a765-4396-9d77-4d7e26b4161a
begin
    directory_download::String = joinpath(pwd(), "download")
    directory_excel::String = joinpath(pwd(), "excel")
end;

# ╔═╡ 7400309f-9473-4a3b-9a50-a85b19f7fba2
md"""
The pages with RFR term structures (recent and previous ones) are scraped.
Results are returned as a dictionary with dates as keys and urls to zip-files
with RFR term structures as values. The date keys are in the format *yyyymmdd*:
for use in lists these dates are converted to a *dd-mm-yyyy* format.
"""

# ╔═╡ 8392fc22-2379-42ce-91a3-e47922f147d6
md"""
Download zip-files with selected RFR term structures, extract the Excel file
containing the term structures, read the relevant cells in the worksheet with
term structures and save the results in an array of SelectedRFR structs.
"""

# ╔═╡ 799bbccf-555d-4852-8f6b-d864fcbb4f60
# Data structure for storing information about a selected RFR term structure.
# rfrDate       : date of selected RFR, format dd-mm-yyyy.
# excelFile     : full path to downloaded (unzipped) Excel file.
# interestRates : vector with interest rates, 150 entries.
struct SelectedRFR
    rfrDate::String
    excelFile::String
    interestRates::Vector{Float64}
end

# ╔═╡ a1270517-c68b-4244-b548-a2b056ec2922
md"""
# Supporting functions
"""

# ╔═╡ 9a36d19d-a5c3-46bf-979b-1c2e2e2d5f52
function get_rfr_urls(page_body::Gumbo.HTMLDocument)::Dict{String, String}
    # Get hyperlinks (HTML <a> tag): gives back a Vector of type Gumbo.HTMLNode.
    # Contains elements of type Gumbo.HTMLElement.
    links_current_rfr = eachmatch(Cascadia.Selector("a"), page_body.root)
    # Get attributes of HTMLElement's: gives back a Vector of Dict's
    attrs_current_rfr = [link.attributes for link in links_current_rfr]
    # Get url's which point to zip-files: select on class key
    urls_zip = [attr["href"] for attr in attrs_current_rfr 
                             if get(attr, "class", "error") == "related-item file-type-zip"]
    valid_urls = Dict{String, String}()
    month_end = ["0131", "0228", "0229", "0331", "0430", "0531", "0630", 
                 "0731", "0831", "0930", "1031", "1130", "1231"]
    for url in urls_zip
        # Before the .zip extension there should always be a number (this excludes
        # special zip files which are also scraped.
        if isdigit(url[end-4])
            for month in month_end
                if contains(url, month)
                    valid_urls[basename(url)[11:18]] = url
                    break
                end
            end
        end
    end
    return valid_urls
end;

# ╔═╡ 0e55dc15-7df0-4655-88a3-96e2972dc536
function scrape_eiopa_website(urls::Vector{String})::Dict{String, String}

    zip_urls = Dict{String, String}()
    for url in urls
        page = HTTP.request("GET", url;verbose=0)
        if page.status != 200
            error("URL could not be scraped.")
        end
        page_body = Gumbo.parsehtml(String(page.body))
        zip_urls = merge(zip_urls, get_rfr_urls(page_body))
    end
    return zip_urls

end;

# ╔═╡ 9a973a7a-bbe4-4e38-ba7c-bc9db682aa1a
function validate_dir(directory::String, create::Bool=true)::Bool
    if isdir(directory)
        return true
    end
    if create
        return isdir(mkpath(directory))
    end
    return false
end;

# ╔═╡ 663c1526-14b0-41ac-8e27-3a6e2d2534eb
function download_zip(url::String, download_dir::String, overwrite::Bool=false)::String
    if !validate_dir(download_dir)
        error("Invalid download directory.")
    end
    download_filename::String = joinpath(download_dir, basename(url))
    if isfile(download_filename) && !overwrite
        # File already present and not to overwritten
        return download_filename
    end
    return Downloads.download(url, download_filename)
end;

# ╔═╡ e492cadc-1e4e-4107-a710-84edb069ec67
"""Returns an array with names of files in the specified zip-file."""
function files_in_zip(zipname::String)::Vector{String}
    files = Vector{String}()
    r = ZipFile.Reader(zipname)
    for file in r.files
        push!(files, file.name)
    end
    close(r)
    return files
end;

# ╔═╡ 5715ebb6-b866-4ddf-8999-db71a93bdc1d
"""Extract Excel file with risk free interest term structure from zip-file
   and return the full path to the uncompressed file."""
function process_zip(zipname::String, process_dir::String)::String
    if !validate_dir(process_dir)
        error("Invalid download directory.")
    end
    filename = ""
    r = ZipFile.Reader(zipname)
    for file in r.files
        if contains(file.name, "_Term_Structures.xlsx")
            filename = joinpath(process_dir, file.name)
            ZipFile.write(filename, ZipFile.read(file))
            break
        end
    end
    close(r)
    return filename
end;

# ╔═╡ ee80e807-c732-4b00-bda1-40190601a2e2
begin
    convert_date_to_ddmmyyyy = (date::String -> string(SubString(date, 7:8), "-",
                                                       SubString(date, 5:6), "-",
                                                       SubString(date, 1:4))::String)
    convert_date_to_yyyymmdd = (date::String -> string(SubString(date, 7:10),
                                                       SubString(date, 4:5),
                                                       SubString(date, 1:2))::String)
end;

# ╔═╡ 6b06aca8-726f-4ed1-a53b-8894cfc51dfd
begin
    # Scrape website: collect all hyperlinks to zip-files with RFR term structures.
    # key: date of RFR term structure in format yyyymmdd, value is the actual url to the zip-file.
    rfr_files::Dict{String, String} = scrape_eiopa_website(rfr_urls)
    # Sort keys of rfr_files and change representation for use in selection list.
    # Representation: yyyymmdd -> dd-mm-yyyy.
    rfr_dates::Vector{String} = reverse([convert_date_to_ddmmyyyy(key) for key in sort(collect(keys(rfr_files)))])
end;

# ╔═╡ d0bbbdae-95e5-4448-ad80-ba7ca22f3178
begin
if download_all
    for key in keys(rfr_files)
        download_zip(rfr_files[key], joinpath(pwd(), "downloads"))
    end
end
rfr_dates_list = copy(rfr_dates)
pushfirst!(rfr_dates_list, "")
md"""
Select dates of risk-free interest rate term structures to be shown in 
the plot below (maximum of 4):\
$(@bind rfr_1 PlutoUI.Select(rfr_dates;default=rfr_dates[1]))
$(@bind rfr_2 PlutoUI.Select(rfr_dates_list;default=rfr_dates_list[1]))
$(@bind rfr_3 PlutoUI.Select(rfr_dates_list;default=rfr_dates_list[1]))
$(@bind rfr_4 PlutoUI.Select(rfr_dates_list;default=rfr_dates_list[1]))
"""
end

# ╔═╡ 1d0e2609-cf22-4169-9e3c-dc735ffc3e87
begin
    # Check which dates have been specified in the selection lists: these RFR term structures
    # will be downloaded and unzipped. There will always be one date present (the first one), 
    # the other ones are optional (""). The selected RFR's are stored as a vector of SelectedRFR's.
    selected_rfrs = Vector{SelectedRFR}()
    for rfr in [rfr_1, rfr_2, rfr_3, rfr_4]
        if rfr != ""
            # A date has been selected
            zip_file = download_zip(rfr_files[convert_date_to_yyyymmdd(rfr)],
                                    directory_download)
            excel_file = process_zip(zip_file, directory_excel)
            excel_data = XLSX.readdata(excel_file, "RFR_spot_no_VA", "C11:C160")
            # Store properties of selection in a SelectedRFR struct. Note that rfr_upper_limit
            # is set by the slider directly above the plot.
            push!(selected_rfrs, SelectedRFR(rfr,
                                             excel_file,
                                             [excel_data[i] for i in 1:1:rfr_upper_limit]))								
        end
    end
end;

# ╔═╡ 7de74ab8-dc5c-411a-995c-d29f67307e5e
begin
    rfr_plot = Plots.plot(1:rfr_upper_limit, 
                          selected_rfrs[1].interestRates,
                          label=selected_rfrs[1].rfrDate, 
                          legend=:topleft,
                          linewidth=3,
                          size=(1000,  600),
                          title="RFR term structures")
    Plots.xlabel!(rfr_plot, "Projection in years", bottom_margin=5Plots.PlotMeasures.mm)
    Plots.ylabel!(rfr_plot, "Risk-free interest rate (%)", left_margin=5Plots.PlotMeasures.mm)
    for i in 2:length(selected_rfrs)
        Plots.plot!(rfr_plot,
                    1:rfr_upper_limit,
                    selected_rfrs[i].interestRates, 
                    label=selected_rfrs[i].rfrDate,
                    linewidth=3)
    end
    Plots.yticks!(rfr_plot, (Plots.yticks(rfr_plot)[1][1], 
                             [string(100*tick, "%") for tick in Plots.yticks(rfr_plot)[1][1]]))
    rfr_plot
end

# ╔═╡ 682e0515-22b8-437f-af62-47a490652e9b
html"""<style>
main {
    max-width: 60%;
}
"""

# ╔═╡ Cell order:
# ╟─282cd541-198e-4419-926d-13bb25826fc3
# ╟─a5bc5425-0972-47cc-be22-9aeabf2345bb
# ╟─849e4709-1dea-4cae-a7ff-34ff43f6a613
# ╟─9e7d2f3d-c45e-427b-bbde-f3d00f55a3b4
# ╟─e77d8d03-c9b0-4ed6-ba64-fd7ab03f98db
# ╟─d0bbbdae-95e5-4448-ad80-ba7ca22f3178
# ╟─f57e9bfc-7b64-4460-ac3c-05573a32859c
# ╟─7de74ab8-dc5c-411a-995c-d29f67307e5e
# ╟─7c527a88-7eb7-4c6c-b14c-c6f22ea0083e
# ╠═0d6afd42-c7ca-11eb-27d6-1db1a07a7e73
# ╟─1322974d-16bd-4fea-853a-335c77d1721a
# ╟─efc29f37-e649-4d5c-beaf-082c7c11f9f3
# ╠═d8399f9d-8984-41d6-970a-cf7d65b2cac1
# ╟─8fcc98e7-1036-4cbd-8ce1-d745d2c50484
# ╠═d8fbcacf-a765-4396-9d77-4d7e26b4161a
# ╟─7400309f-9473-4a3b-9a50-a85b19f7fba2
# ╠═6b06aca8-726f-4ed1-a53b-8894cfc51dfd
# ╟─8392fc22-2379-42ce-91a3-e47922f147d6
# ╠═799bbccf-555d-4852-8f6b-d864fcbb4f60
# ╠═1d0e2609-cf22-4169-9e3c-dc735ffc3e87
# ╟─a1270517-c68b-4244-b548-a2b056ec2922
# ╠═0e55dc15-7df0-4655-88a3-96e2972dc536
# ╠═9a36d19d-a5c3-46bf-979b-1c2e2e2d5f52
# ╠═9a973a7a-bbe4-4e38-ba7c-bc9db682aa1a
# ╠═663c1526-14b0-41ac-8e27-3a6e2d2534eb
# ╠═e492cadc-1e4e-4107-a710-84edb069ec67
# ╠═5715ebb6-b866-4ddf-8999-db71a93bdc1d
# ╠═ee80e807-c732-4b00-bda1-40190601a2e2
# ╟─682e0515-22b8-437f-af62-47a490652e9b
