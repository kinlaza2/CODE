
        echo "------------ CHECKING   ConsumerCount   --------------"

for i in eventsQueue mtsQueue mosQueue  customerCareQueue moPreprocessQueue ticketQueue creditQueue  datasyncQueue  
do
 	for j in  ConsumerCount
	do
		./check_HQ $i $j
	done
	sleep 2

done

        echo "------------ CHECKING   MessageCount   --------------"

for k in eventsEXP eventsDLQ mtsEXP mtsDLQ customerCareEXP customerCareDLQ mosEXP mosDLQ moPreprocessQueueEXP   moPreprocessQueueDLQ ticketQueueEXP ticketQueueDLQ eventsQueue mtsQueue mosQueue  customerCareQueue moPreprocessQueue ticketQueue creditQueue  creditDLQ creditEXP datasyncQueue  datasyncDLQ datasyncEXP

	do
		./check_HQ $k MessageCount
	done
	sleep 2

        echo "------------ CHECKING   MessagesAdded   --------------"
for w in eventsQueue mtsQueue mosQueue  customerCareQueue moPreprocessQueue ticketQueue creditQueue  datasyncQueue  
	do
		./check_HQ $w MessagesAdded  5
	done




