% plot_resistivity.m
% Read Excel file, extract column 13 (apparent resistivity) and column 14 (error),
% and create the plot.
% Left y-axis: Apparent Resistivity (blue), Right y-axis: Error (red),
% x-axis: measurement point / distance
%
% Usage: Run the script directly → file selection dialog will appear.
%        Or hardcode the 'fullpath' variable below and comment out the uigetfile part.

%% Select file (comment out the next line and set fullpath if you want to hardcode the path)
[filename, pathname] = uigetfile({'*.xlsx;*.xls','Excel Files (*.xlsx, *.xls)'; '*.*','All Files'}, ...
                                 'Select Excel File');
if isequal(filename, 0)
    error('No file selected. Script terminated.');
end
fullpath = fullfile(pathname, filename);

%% Read data (prefer readtable to handle headers properly)
try
    T = readtable(fullpath);
    if size(T, 2) < 15
        error('Table has fewer than 15 columns.');
    end
    col13 = T{:, 13};
    col14 = T{:, 14};
    col15 = T{:, 15};
catch ME
    % If readtable fails, try readmatrix
    try
        M = readmatrix(fullpath);
        if size(M, 2) < 15
            rethrow(ME);
        end
        col13 = M(:, 13);
        col14 = M(:, 14);
        col15 = M(:, 15);
    catch
        error('Failed to read file or insufficient columns (need at least 15). Original error: %s', ME.message);
    end
end

%% Convert to column vectors and clean non-numeric values
col13 = col13(:);
col14 = col14(:);
col15 = col15(:);

N = max(length(col13), length(col14));

% Pad shorter column with NaN if lengths differ
if length(col13) < N
    col13(end+1:N) = NaN;
end
if length(col14) < N
    col14(end+1:N) = NaN;
end

% Prefer rows where both resistivity and error are valid
validBoth = ~isnan(col13) & ~isnan(col14);
if any(validBoth)
    rho = col13(validBoth);
    err = col14(validBoth);
    x = 1:length(rho);
else
    % Fall back to resistivity column only
    valid13 = ~isnan(col13);
    if ~any(valid13)
        error('No valid values found in column 13 (apparent resistivity). Cannot plot.');
    end
    rho = col13(valid13);
    err = col14(valid13);   % may contain NaN → will show gaps in plot
    x = 1:length(rho);
end

%% Plotting
% --- Font and style settings ---
font_name = 'Times New Roman';
fs        = 24;   % tick label font size
label_fs  = 26;   % axis label font size
legend_fs = 24;   % legend font size
line_w    = 1.5;  % line width

figure('Name', 'Comparison of Concurrent and Single Mode Measurements', ...
       'NumberTitle', 'off', 'Color', 'w');

% =======================================================
% Left axis: Apparent Resistivity
% =======================================================
yyaxis left;

% Plot: blue circles for concurrent, green triangles for single
p1 = plot(x, col13(1:length(x)), '-o', 'LineWidth', line_w, ...
          'MarkerSize', 8, 'MarkerFaceColor', 'b');
hold on;
p2 = plot(x, col14(1:length(x)), '-^', 'LineWidth', line_w, ...
          'MarkerSize', 8, 'MarkerFaceColor', 'g');

% Labels
ylabel('Apparent Resistivity (\Omega \cdot m)', ...
       'FontName', font_name, 'FontWeight', 'bold', 'FontSize', label_fs);
xlabel('Distance (m)', ...
       'FontName', font_name, 'FontWeight', 'bold', 'FontSize', label_fs);

ax = gca;
ax.YColor = [0 0 0.8];  % dark blue

% --- Left axis limits and ticks ---
ylim([0, 300]);
yticks(0:50:300);
% -------------------------------------------
grid on;

% =======================================================
% Right axis: Apparent Resistivity Error
% =======================================================
yyaxis right;

% Plot: red squares for error
p3 = plot(x, col15(1:length(x)), '-s', 'LineWidth', line_w, ...
          'MarkerSize', 8, 'MarkerFaceColor', 'r');

% Labels
ylabel('Apparent Resistivity Error (%)', ...
       'FontName', font_name, 'FontWeight', 'bold', 'FontSize', label_fs);

ax.YColor = [0.8 0 0];  % dark red

% --- Right axis limits and ticks ---
ylim([0, 5]);
yticks(0:0.5:5);
% -------------------------------------------

% =======================================================
% Global font settings
% =======================================================
set(ax, 'FontName', font_name, 'FontSize', fs, ...
        'FontWeight', 'bold', 'LineWidth', 1.2);

% =======================================================
% Legend (update frequencies manually when needed)
% =======================================================
lgd = legend([p1, p2, p3], ...
    'Concurrent Mode (1 Hz)', ...
    'Single Mode (1 Hz)', ...
    'Apparent Resistivity Error');

set(lgd, 'FontName', font_name, 'FontSize', legend_fs, ...
         'Location', 'northeast', 'Box', 'on');

%% Finish
disp('Plot completed.');