#!/usr/bin/env bash


VM_PATH=/home/lnikos/VM__MACHINES
VM_GUEST_FILES_PATHS="/home/lnikos/VM__MACHINES/WIN7_SUPPORT/Windows_7  /home/lnikos/VM__MACHINES/CEN_SUP/CEN_SUP"
SCRIPT_PATH=/root/STARTUP_SCRIPTS


############# Check if any process run for the specific VMs and stop them if they run
 for i in  `echo $VM_GUEST_FILES_PATHS`
 do
         # check if we have any VM Runs on the specific paths
         VM_RUN=`ps -ef | grep vmware | grep $i | wc -l`

         ## TO BE ADDED CHECK IF $VM_RUN isnumeric       !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

         #a is the path to corresponding vmx file
         a=`echo -e $i/*.vmx`
         if [ $VM_RUN -eq 0 ]
         then
                 echo -e "VM $a is not running"
                 /bin/rm -rf   $i/*.lck
         else
                 # if VM is up then restarted.
                 /usr/bin/vmrun stop  $i/*.vmx
                 sleep 2;
                 /bin/rm -rf   $i/*.lck
                 echo -e "VM $a was STOPPED"
         fi
 done


############# START REQUESTED VMs #####################################################################
for i in  `echo $VM_GUEST_FILES_PATHS`
do

/bin/rm -rf   $i/*.lck
/usr/bin/vmrun -T ws start  $i/*.vmx nogui

#/bin/rm -rf /home/lnikos/VM__MACHINES/RHE_5_2_NIK1_EXTERNAL_ACCESS/*.lck
#/usr/bin/vmrun  -T ws start /home/lnikos/VM__MACHINES/RHE_5_2_NIK1_EXTERNAL_ACCESS/Red\ Hat\ Enterprise\ Linux\ 5.vmx nogui

done
