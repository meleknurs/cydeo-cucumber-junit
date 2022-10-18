Feature: Examples of the Cucumber data table implementations


  Scenario: List of fruits I like
    Then user should see fruits I like
      | kiwi     |
      | banana   |
      | cucumber |
      | orange   |
      | apple    |
      | avacado  |
