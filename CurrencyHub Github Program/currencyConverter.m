function currencyConverter
    % Load the spreadsheet data
    filename = 'Currency-Converter-Power-Query.xlsx';
    data = readtable(filename, 'VariableNamingRule', 'preserve');
    
    % Remove rows with empty or NaN values
    data = data(~any(ismissing(data), 2), :);
    
    % Debug: Display cleaned data
    disp('Cleaned data:');
    disp(data);
    
    % Create the main figure
    fig = uifigure('Name', 'Currency Converter', 'Position', [100, 100, 400, 600]); % Increased height for visibility
    
    % Create a dropdown for selecting the base currency
    uilabel(fig, 'Position', [50, 450, 100, 30], 'Text', 'From:');
    baseCurrencyDropdown = uidropdown(fig, 'Position', [150, 450, 200, 30]);
    
    % Create a dropdown for selecting the target currency
    uilabel(fig, 'Position', [50, 410, 100, 30], 'Text', 'To:');
    targetCurrencyDropdown = uidropdown(fig, 'Position', [150, 410, 200, 30]);
    
    % Create an edit field for the amount
    uilabel(fig, 'Position', [50, 370, 100, 30], 'Text', 'Amount:');
    amountField = uieditfield(fig, 'numeric', 'Position', [150, 370, 200, 30]);
    
    % Create a button for conversion
    convertButton = uibutton(fig, 'Position', [150, 320, 200, 30], 'Text', 'Convert', ...
        'ButtonPushedFcn', @(convertButton, event) convertCurrency(baseCurrencyDropdown, targetCurrencyDropdown, amountField, data, fig));
    
    % Create an axes for plotting
    ax = uiaxes(fig, 'Position', [50, 260, 300, 60]); % Adjusted position to fit all elements
    title(ax, 'Conversion Rate Over Time');
    xlabel(ax, 'Time');
    ylabel(ax, 'Conversion Rate');
    
    % Populate the dropdowns with the currency names
    currencies = unique([data.('From:'); data.('To:')]); % Unique currencies from both columns
    baseCurrencyDropdown.Items = currencies;
    targetCurrencyDropdown.Items = currencies;
    
    % Debug: Display dropdown items
    disp('Dropdown items:');
    disp(baseCurrencyDropdown.Items);
    disp(targetCurrencyDropdown.Items);
    
    % Input fields for adding new conversion
    uilabel(fig, 'Position', [50, 200, 100, 30], 'Text', 'New From:');
    newFromField = uieditfield(fig, 'text', 'Position', [150, 200, 200, 30]);
   
    uilabel(fig, 'Position', [50, 160, 100, 30], 'Text', 'New To:');
    newToField = uieditfield(fig, 'text', 'Position', [150, 160, 200, 30]);
   
    uilabel(fig, 'Position', [50, 120, 100, 30], 'Text', 'New Amount:');
    newAmountField = uieditfield(fig, 'numeric', 'Position', [150, 120, 200, 30]);
   
    uilabel(fig, 'Position', [50, 80, 100, 30], 'Text', 'New Result:');
    newResultField = uieditfield(fig, 'numeric', 'Position', [150, 80, 200, 30]);
   
    % Move the Add Conversion button to be visible
    addButton = uibutton(fig, 'Position', [150, 20, 200, 30], 'Text', 'Add Conversion', ...
       'ButtonPushedFcn', @(addButton, event) addConversion(newFromField, newToField, newAmountField, newResultField, data, baseCurrencyDropdown, targetCurrencyDropdown));
end

function convertCurrency(baseCurrencyDropdown, targetCurrencyDropdown, amountField, data, fig)
    try
        % Get the selected currencies and amount
        baseCurrency = baseCurrencyDropdown.Value;
        targetCurrency = targetCurrencyDropdown.Value;
        amount = amountField.Value;
        
        % Debug: Display selected currencies and amount
        disp(['Selected base currency: ' baseCurrency]);
        disp(['Selected target currency: ' targetCurrency]);
        disp(['Amount: ' num2str(amount)]);
        
        % Check if a direct conversion rate is available
        baseToTargetRate = data(strcmp(data.('From:'), baseCurrency) & strcmp(data.('To:'), targetCurrency), :);
        
        if ~isempty(baseToTargetRate)
            % Direct conversion available
            amt = baseToTargetRate.('Amt:');
            result = baseToTargetRate.('Result:');
            conversionRate = result / amt;
            convertedAmount = amount * conversionRate;
            uialert(fig, ['Converted Amount: ' num2str(convertedAmount)], 'Conversion Result');
        else
            % Indirect conversion using US Dollar as an intermediary
            intermediary = 'US Dollar';
            
            % Convert base currency to intermediary
            baseToIntermediary = data(strcmp(data.('From:'), baseCurrency) & strcmp(data.('To:'), intermediary), :);
            if isempty(baseToIntermediary)
                uialert(fig, 'Direct conversion rate not found and cannot use intermediary.', 'Error');
                return;
            end
            baseToIntermediaryRate = baseToIntermediary.('Result:') / baseToIntermediary.('Amt:');
            
            % Convert intermediary to target currency
            intermediaryToTarget = data(strcmp(data.('From:'), intermediary) & strcmp(data.('To:'), targetCurrency), :);
            if isempty(intermediaryToTarget)
                uialert(fig, 'Indirect conversion rate not found for the selected currencies.', 'Error');
                return;
            end
            intermediaryToTargetRate = intermediaryToTarget.('Result:') / intermediaryToTarget.('Amt:');
            
            % Calculate the converted amount using intermediary rates
            convertedAmount = amount * baseToIntermediaryRate * intermediaryToTargetRate;
            uialert(fig, ['Converted Amount: ' num2str(convertedAmount)], 'Conversion Result');
        end
        
        % Plot example data for the conversion rate over time
        time = 1:10; % Example time data
        conversionRates = randn(size(time)) * 0.05; % Example conversion rates with some noise
        
        % Plot the conversion rates
        ax = findobj(fig, 'Type', 'Axes');
        plot(ax, time, conversionRates);
        
    catch ME
        uialert(fig, ['An error occurred: ' ME.message], 'Error');
    end
end

function addConversion(newFromField, newToField, newAmountField, newResultField, data, baseCurrencyDropdown, targetCurrencyDropdown)
    try
        % Get the new conversion data from input fields
        newFrom = newFromField.Value;
        newTo = newToField.Value;
        newAmount = newAmountField.Value;
        newResult = newResultField.Value;
        
        % Create a new row for the table
        newRow = table({newFrom}, {newTo}, newAmount, newResult, 'VariableNames', {'From:', 'To:', 'Amt:', 'Result:'});
        
        % Append the new row to the existing data
        data = [data; newRow];
        
        % Save the updated table back to the Excel file
        filename = 'Currency-Converter-Power-Query.xlsx';
        writetable(data, filename, 'Sheet', 1);
        
        % Update the dropdowns with the new currency options
        currencies = unique([data.('From:'); data.('To:')]);
        baseCurrencyDropdown.Items = currencies;
        targetCurrencyDropdown.Items = currencies;
        
        % Debug: Display updated data
        disp('Updated data:');
        disp(data);
    catch ME
        % Display an error message if something goes wrong
        uialert(fig, ['An error occurred: ' ME.message], 'Error');
    end
end
