library(tidyverse)
library(xlsx)

data <- read.xlsx("./data.xlsx", sheetName = "Lhuentse")

grouped_data <- data %>%
  select(Gewog, Occupancy.Label, MBT.Label, Total.Floor.Area..m2.) %>% 
  group_by(MBT.Label, Gewog, Occupancy.Label) %>% 
  summarise(count = sum(Total.Floor.Area..m2.)) %>% 
  ungroup()

list_of_gewogs <- grouped_data %>% 
  distinct(Gewog)

complete_grid <- expand.grid(
  Gewog = list_of_gewogs$Gewog,
  MBT.Label = unique(grouped_data$MBT.Label),
  Occupancy.Label = unique(grouped_data$Occupancy.Label)
)

complete_data <- complete_grid %>%
  left_join(grouped_data, by = c("Gewog", "MBT.Label", "Occupancy.Label")) %>%
  replace_na(list(count = 0))

Yangtse_wide_data <- complete_data %>%
  pivot_wider(names_from = Occupancy.Label, values_from = count, values_fill = list(count = 0))

Yangtse_wide_data <- list_of_gewogs %>%
  left_join(Yangtse_wide_data, by = "Gewog") %>%
  replace_na(list(MBT.Label = unique(grouped_data$MBT.Label)))

desired_order <- c("Gewog", "MBT.Label", "RES-SU", "RES-CU", "COM", "IND", "HOS", "EDU", "OFF", "REL", "STOR", "ASM")

Yangtse_wide_data <- Yangtse_wide_data[, desired_order]
  
write.xlsx(Yangtse_wide_data, file = "Yangtse_wide_data.xlsx")

